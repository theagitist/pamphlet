import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { parsePptx, resolveColors, resolveInheritedStyles, extractImages } from '../converter/pptxParser.js';
import { applyTypographyScaling } from '../converter/styleMapper.js';
import { optimizeImages } from '../converter/imageOptimizer.js';
import { writeDocx } from '../converter/docxWriter.js';

const PHASES = ['Parsing', 'Extraction', 'Formatting', 'Writing'];

export async function convert(uploadPath, downloadPath, onProgress) {
  const report = (phase) => {
    if (onProgress) onProgress({ phase, phaseIndex: PHASES.indexOf(phase), total: PHASES.length });
  };

  // Phase 1: Parsing
  report('Parsing');
  const buffer = await readFile(uploadPath);
  const parsed = await parsePptx(buffer);

  // Phase 2: Extraction
  report('Extraction');
  await extractImages(buffer, parsed);
  await optimizeImages(parsed);

  // Phase 3: Formatting
  report('Formatting');
  resolveInheritedStyles(parsed);
  resolveColors(parsed);
  applyTypographyScaling(parsed);

  // Phase 4: Writing
  report('Writing');
  const docxBuffer = await writeDocx(parsed);
  await mkdir(path.dirname(downloadPath), { recursive: true });
  await writeFile(downloadPath, docxBuffer);

  return {
    slides: parsed.slides.length,
    unsupportedObjects: parsed.unsupportedObjects,
  };
}
