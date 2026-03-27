import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { parsePptx, resolveColors, resolveInheritedStyles, extractImages } from '../converter/pptxParser.js';
import { applyTypographyScaling } from '../converter/styleMapper.js';
import { optimizeImages } from '../converter/imageOptimizer.js';
import { renderSmartArt } from '../converter/smartArtRenderer.js';
import { writeDocx } from '../converter/docxWriter.js';

const PHASES = ['Parsing', 'Extraction', 'Formatting', 'Writing'];

export async function convert(uploadPath, downloadPath, onProgress, { workDir } = {}) {
  const report = (phase) => {
    if (onProgress) onProgress({ phase, phaseIndex: PHASES.indexOf(phase), total: PHASES.length });
  };

  // Phase 1: Parsing
  report('Parsing');
  const buffer = await readFile(uploadPath);
  const parsed = await parsePptx(buffer);

  // Phase 2: Extraction — images written to disk in workDir, not held in RAM
  report('Extraction');
  if (workDir) {
    await mkdir(workDir, { recursive: true });
    await extractImages(parsed.zip, parsed, workDir);
    // Render SmartArt diagrams as images via LibreOffice
    await renderSmartArt(parsed, uploadPath, workDir);
  } else {
    // Fallback for tests without workDir — extract to buffers in memory
    await extractImages(parsed.zip, parsed, null);
  }
  await optimizeImages(parsed);

  // Release the zip object to free memory
  parsed.zip = null;

  // Phase 3: Formatting
  report('Formatting');
  resolveInheritedStyles(parsed);
  resolveColors(parsed);
  applyTypographyScaling(parsed);

  // Phase 4: Writing — reads each image from disk one at a time
  report('Writing');
  const docxBuffer = await writeDocx(parsed);
  await mkdir(path.dirname(downloadPath), { recursive: true });
  await writeFile(downloadPath, docxBuffer);

  return {
    slides: parsed.slides.length,
    unsupportedObjects: parsed.unsupportedObjects,
  };
}
