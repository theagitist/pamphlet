// Renders SmartArt slides as images using LibreOffice + pdftoppm.
//
// Strategy:
// 1. Find all slides containing SmartArt elements
// 2. Create a temp copy of the PPTX with non-SmartArt elements stripped from those slides
// 3. Convert the cleaned PPTX to PDF via LibreOffice headless
// 4. Extract only the SmartArt slides as PNGs via pdftoppm
// 5. Crop each PNG to the SmartArt bounding box using sharp
// 6. Replace SmartArt elements with image elements

import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import fs from 'node:fs';
import path from 'node:path';
import sharp from 'sharp';
import JSZip from 'jszip';

const execFileAsync = promisify(execFile);

// Standard slide dimensions in EMU (widescreen 13.33" x 7.5")
const SLIDE_WIDTH_EMU = 12192000;
const SLIDE_HEIGHT_EMU = 6858000;

const RENDER_DPI = 200;

/**
 * Create a copy of the PPTX where SmartArt slides have all non-SmartArt
 * elements removed, so the rendered image is clean.
 */
async function createCleanPptx(uploadPath, smartArtSlideIndices, pdfDir) {
  const buf = await fs.promises.readFile(uploadPath);
  const zip = await JSZip.loadAsync(buf);

  for (const slideIdx of smartArtSlideIndices) {
    const slidePath = `ppt/slides/slide${slideIdx}.xml`;
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;

    let xml = await slideFile.async('text');

    // Remove background element (p:bg) so the slide has a white/transparent background
    xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/g, '');

    // Remove all p:sp (shapes/text), p:pic (images), p:grpSp (groups) from the shape tree,
    // but keep p:graphicFrame (which contains SmartArt/tables/charts).
    // We do this by matching and removing specific elements within <p:spTree>.
    // The shape tree starts after <p:grpSpPr>...</p:grpSpPr>
    xml = xml.replace(
      /(<p:spTree>[\s\S]*?<\/p:grpSpPr>)([\s\S]*?)(<\/p:spTree>)/g,
      (match, prefix, content, suffix) => {
        // Keep only p:graphicFrame elements (SmartArt, tables, etc.)
        const kept = [];
        // Match each top-level element in the shape tree content
        const elementPattern = /<p:(sp|pic|grpSp|graphicFrame)\b[\s\S]*?<\/p:\1>/g;
        let m;
        while ((m = elementPattern.exec(content)) !== null) {
          if (m[1] === 'graphicFrame') {
            kept.push(m[0]);
          }
        }
        return prefix + kept.join('') + suffix;
      }
    );

    zip.file(slidePath, xml);
  }

  const cleanPath = path.join(pdfDir, 'clean.pptx');
  const outBuf = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.promises.writeFile(cleanPath, outBuf);
  return cleanPath;
}

/**
 * Render SmartArt elements as images.
 * Modifies parsedData in place, converting smartart elements to image elements.
 */
export async function renderSmartArt(parsedData, uploadPath, workDir) {
  // Collect slide indices (1-based) that have SmartArt
  const smartArtSlides = new Map(); // slideIndex → [element refs]

  for (const slide of parsedData.slides) {
    for (const el of flatElements(slide.elements)) {
      if (el.type === 'smartart') {
        const idx = slide.index;
        if (!smartArtSlides.has(idx)) smartArtSlides.set(idx, []);
        smartArtSlides.get(idx).push(el);
      }
    }
  }

  if (smartArtSlides.size === 0) return;

  const pdfDir = path.join(workDir, '_smartart');
  await fs.promises.mkdir(pdfDir, { recursive: true });

  try {
    // Create a cleaned PPTX with only SmartArt on affected slides
    const cleanPptx = await createCleanPptx(
      uploadPath,
      [...smartArtSlides.keys()],
      pdfDir,
    );

    // Convert cleaned PPTX → PDF using LibreOffice headless
    await execFileAsync('libreoffice', [
      '--headless', '--convert-to', 'pdf',
      '--outdir', pdfDir,
      cleanPptx,
    ], { timeout: 60000 });

    // Find the generated PDF
    const pdfFiles = await fs.promises.readdir(pdfDir);
    const pdf = pdfFiles.find(f => f.endsWith('.pdf'));
    if (!pdf) return;
    const actualPdf = path.join(pdfDir, pdf);

    // For each SmartArt slide, extract as PNG and crop to SmartArt bounds
    for (const [slideIdx, elements] of smartArtSlides) {
      const slidePrefix = path.join(pdfDir, `slide_${slideIdx}`);

      await execFileAsync('pdftoppm', [
        '-png', '-r', String(RENDER_DPI),
        '-f', String(slideIdx), '-l', String(slideIdx),
        actualPdf, slidePrefix,
      ], { timeout: 30000 });

      // pdftoppm names output: slidePrefix-NN.png
      const pngFiles = await fs.promises.readdir(pdfDir);
      const slidePng = pngFiles.find(f =>
        f.startsWith(`slide_${slideIdx}-`) && f.endsWith('.png')
      ) || pngFiles.find(f =>
        f.startsWith(`slide_${slideIdx}`) && f.endsWith('.png')
      );

      if (!slidePng) continue;
      const fullSlidePath = path.join(pdfDir, slidePng);

      // Get actual rendered image dimensions
      const metadata = await sharp(fullSlidePath).metadata();
      const imgW = metadata.width;
      const imgH = metadata.height;

      // Crop and save for each SmartArt element on this slide
      for (let i = 0; i < elements.length; i++) {
        const el = elements[i];
        const pos = el.position || {};

        // Convert EMU position to pixel coordinates with padding
        const padX = Math.round(imgW * 0.02);
        const padY = Math.round(imgH * 0.03);

        const rawX = Math.round((pos.x / SLIDE_WIDTH_EMU) * imgW);
        const rawY = Math.round((pos.y / SLIDE_HEIGHT_EMU) * imgH);
        const rawW = Math.round((pos.cx / SLIDE_WIDTH_EMU) * imgW);
        const rawH = Math.round((pos.cy / SLIDE_HEIGHT_EMU) * imgH);

        const cropX = Math.max(0, rawX - padX);
        const cropY = Math.max(0, rawY - padY);
        const cropW = Math.min(imgW - cropX, rawW + padX * 2);
        const cropH = Math.min(imgH - cropY, rawH + padY * 2);

        if (cropW <= 0 || cropH <= 0) {
          el.type = 'unsupported';
          el.description = 'SmartArt (render failed)';
          continue;
        }

        const croppedPath = path.join(workDir, `smartart_s${slideIdx}_${i}.png`);
        await sharp(fullSlidePath)
          .extract({ left: cropX, top: cropY, width: cropW, height: cropH })
          .png({ quality: 90 })
          .toFile(croppedPath);

        // Scale down to fit well in a Word document
        const scaleFactor = 0.55;
        el.type = 'image';
        el.imageFile = croppedPath;
        el.imageBuffer = null;
        el.name = el.description || 'SmartArt';
        el.altText = null;
        el.width = Math.round(pos.cx * scaleFactor);
        el.height = Math.round(pos.cy * scaleFactor);

        // Clean up SmartArt-specific fields
        delete el.diagramRels;
        delete el.slideRels;
        delete el.description;
      }
    }
  } catch (err) {
    console.error('SmartArt rendering failed:', err.message);
    // Fall back: mark remaining smartart elements as unsupported
    for (const [, elements] of smartArtSlides) {
      for (const el of elements) {
        if (el.type === 'smartart') {
          el.type = 'unsupported';
          el.description = 'SmartArt (render unavailable)';
          delete el.diagramRels;
          delete el.slideRels;
        }
      }
    }
  }
}

function flatElements(elements) {
  const result = [];
  for (const el of elements) {
    result.push(el);
    if (el.type === 'group' && el.elements) {
      result.push(...flatElements(el.elements));
    }
  }
  return result;
}
