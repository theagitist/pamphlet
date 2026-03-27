import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { parsePptx, extractImages } from '../../src/server/converter/pptxParser.js';
import { optimizeImageFile } from '../../src/server/converter/imageOptimizer.js';
import { writeFileSync, readFileSync, unlinkSync, existsSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { buildTypographyScaler } from '../../src/server/converter/styleMapper.js';
import {
  createSession, getSession, setUploaded, setStatus,
  purgeSession, startExpiryTimer, sessionCount,
} from '../../src/server/services/sessionManager.js';

// ── Parser edge cases ──────────────────────────────────────────

describe('parsePptx edge cases', () => {
  it('throws on non-ZIP buffer', async () => {
    const garbage = Buffer.from('this is not a zip file');
    await expect(parsePptx(garbage)).rejects.toThrow('Invalid or corrupted');
  });

  it('throws on ZIP without presentation.xml', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'not a pptx');
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    await expect(parsePptx(buf)).rejects.toThrow('missing ppt/presentation.xml');
  });

  it('throws on ZIP without presentation relationships', async () => {
    const zip = new JSZip();
    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldIdLst/>
      </p:presentation>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    await expect(parsePptx(buf)).rejects.toThrow('missing presentation relationships');
  });

  it('handles empty presentation (no slides) gracefully', async () => {
    const zip = new JSZip();
    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldIdLst/>
      </p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      </Relationships>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const result = await parsePptx(buf);
    expect(result.slides).toHaveLength(0);
  });

  it('handles empty buffer', async () => {
    await expect(parsePptx(Buffer.alloc(0))).rejects.toThrow();
  });
});

// ── Path traversal prevention ──────────────────────────────────

describe('extractImages path safety', () => {
  it('rejects image paths that escape ppt/ directory', async () => {
    const zip = new JSZip();
    zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
      </p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
      </Relationships>`);
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld><p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
          <p:pic>
            <p:nvPicPr><p:cNvPr id="2" name="evil"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
            <p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
          </p:pic>
        </p:spTree></p:cSld>
      </p:sld>`);
    zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../../etc/passwd"/>
      </Relationships>`);
    // Put a file at the traversal target to see if it gets read
    zip.file('etc/passwd', 'root:x:0:0');

    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const parsed = await parsePptx(buf);
    await extractImages(buf, parsed);

    const img = parsed.slides[0]?.elements.find(e => e.type === 'image');
    expect(img.imageBuffer).toBeNull();
  });
});

// ── Image optimizer edge cases ─────────────────────────────────

describe('optimizeImageFile edge cases', () => {
  const tmpDir = '/tmp/pamphlet-test-img-' + Date.now();

  it('handles nonexistent file gracefully', async () => {
    await expect(optimizeImageFile('/tmp/nonexistent-file.png')).resolves.toBeUndefined();
  });

  it('handles null file path gracefully', async () => {
    await expect(optimizeImageFile(null)).resolves.toBeUndefined();
  });

  it('keeps corrupt file untouched', async () => {
    mkdirSync(tmpDir, { recursive: true });
    const filePath = join(tmpDir, 'corrupt.png');
    writeFileSync(filePath, 'not an image at all');
    await optimizeImageFile(filePath);
    expect(readFileSync(filePath, 'utf8')).toBe('not an image at all');
    unlinkSync(filePath);
  });
});

// ── Typography scaler edge cases ───────────────────────────────

describe('buildTypographyScaler edge cases', () => {
  it('returns identity function for empty slides', () => {
    const scale = buildTypographyScaler({ slides: [] });
    expect(scale(2400)).toBe(2400);
  });

  it('returns identity function for slides with no text', () => {
    const scale = buildTypographyScaler({
      slides: [{ elements: [{ type: 'image' }] }],
    });
    expect(scale(2400)).toBe(2400);
  });

  it('scales all-same-size text to body cap', () => {
    const scale = buildTypographyScaler({
      slides: [{
        elements: [{
          type: 'shape',
          paragraphs: [{
            runs: [
              { type: 'text', props: { size: 2400 } },
              { type: 'text', props: { size: 2400 } },
            ],
          }],
        }],
      }],
    });
    expect(scale(2400)).toBe(1400); // MAX_BODY_PT = 14pt = 1400
  });
});

// ── Session manager edge cases ─────────────────────────────────

describe('sessionManager edge cases', () => {
  it('creates session with UUID', () => {
    const session = createSession();
    expect(session.id).toMatch(/^[0-9a-f-]{36}$/);
    expect(session.status).toBe('created');
    purgeSession(session.id); // cleanup
  });

  it('returns null for nonexistent session', () => {
    expect(getSession('nonexistent-id')).toBeNull();
  });

  it('purging nonexistent session is a no-op', () => {
    expect(() => purgeSession('nonexistent-id')).not.toThrow();
  });

  it('purging same session twice is safe', () => {
    const session = createSession();
    purgeSession(session.id);
    expect(() => purgeSession(session.id)).not.toThrow();
    expect(getSession(session.id)).toBeNull();
  });

  it('setStatus returns null for nonexistent session', () => {
    expect(setStatus('nonexistent', 'processing')).toBeNull();
  });

  it('session is removed after purge', () => {
    const session = createSession();
    const id = session.id;
    purgeSession(id);
    expect(getSession(id)).toBeNull();
  });

  it('expiry timer cleans up session', async () => {
    const session = createSession();
    const id = session.id;
    startExpiryTimer(id, 0.001); // ~60ms
    await new Promise(r => setTimeout(r, 150));
    expect(getSession(id)).toBeNull();
  });
});
