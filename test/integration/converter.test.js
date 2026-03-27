import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync, unlinkSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { convert } from '../../src/server/services/converter.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = resolve(__dirname, '../fixtures/basic.pptx');
const outputPath = resolve(__dirname, '../fixtures/converter-test-output.docx');

describe('converter end-to-end', () => {
  it('converts basic.pptx to docx with all 4 phases', async () => {
    const phases = [];
    const result = await convert(fixturePath, outputPath, ({ phase }) => {
      phases.push(phase);
    });

    expect(phases).toEqual(['Parsing', 'Extraction', 'Formatting', 'Writing']);
    expect(result.slides).toBe(2);
    expect(result.unsupportedObjects).toHaveLength(1);
    expect(existsSync(outputPath)).toBe(true);
  });

  it('output docx contains all source text verbatim', async () => {
    const buf = readFileSync(outputPath);
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file('word/document.xml').async('text');

    // Extract all w:t text nodes
    const textMatches = docXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const allText = textMatches.map(m => m.replace(/<[^>]+>/g, '')).join('');

    // Verify all source text present
    expect(allText).toContain('Welcome to Pamphlet');
    expect(allText).toContain('This is ');
    expect(allText).toContain('bold');
    expect(allText).toContain('italic blue');
    expect(allText).toContain(' text.');
    expect(allText).toContain('Underlined scheme-colored text');
    expect(allText).toContain('Footer note: Confidential');
    expect(allText).toContain('Grouped text A');
    expect(allText).toContain('Grouped text B');
  });

  it('output docx uses no Heading styles', async () => {
    const buf = readFileSync(outputPath);
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file('word/document.xml').async('text');

    expect(docXml).not.toMatch(/Heading/i);
  });

  it('output docx has direct formatting (colors, fonts, bold)', async () => {
    const buf = readFileSync(outputPath);
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file('word/document.xml').async('text');

    // Title: Arial, red, bold
    expect(docXml).toContain('w:val="FF0000"');
    expect(docXml).toContain('w:ascii="Arial"');

    // Italic blue
    expect(docXml).toContain('w:val="0000FF"');
    expect(docXml).toContain('<w:i/>');

    // Scheme color resolved: accent1 → 4472C4
    expect(docXml).toContain('w:val="4472C4"');

    // Unsupported placeholder
    expect(docXml).toContain('[Object not supported: SmartArt');
  });

  // Cleanup
  it('cleanup output file', () => {
    if (existsSync(outputPath)) unlinkSync(outputPath);
  });
});
