import sharp from 'sharp';
import fs from 'node:fs';
import path from 'node:path';

// Max width for images in a Word document (letter page ~6.5" printable at 150 DPI)
const MAX_WIDTH_PX = 975;
const MAX_HEIGHT_PX = 1200;
const JPEG_QUALITY = 80;
const PNG_COMPRESSION = 9;

// Optimize an image file on disk in-place.
// Reads from filePath, writes optimized version back to the same path.
export async function optimizeImageFile(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return;

  try {
    const image = sharp(filePath);
    const metadata = await image.metadata();

    if (!metadata.width || !metadata.height) return;

    const originalSize = fs.statSync(filePath).size;

    let pipeline = image;

    // Resize if larger than max dimensions
    if (metadata.width > MAX_WIDTH_PX || metadata.height > MAX_HEIGHT_PX) {
      pipeline = pipeline.resize(MAX_WIDTH_PX, MAX_HEIGHT_PX, {
        fit: 'inside',
        withoutEnlargement: true,
      });
    }

    // Choose output format
    const isJpeg = metadata.format === 'jpeg' || metadata.format === 'jpg';
    const isPng = metadata.format === 'png';
    const isLargeImage = originalSize > 100 * 1024;

    // Write to a temp file, then replace original only if smaller
    const tmpPath = filePath + '.opt';
    if (isJpeg) {
      await pipeline.jpeg({ quality: JPEG_QUALITY, mozjpeg: true }).toFile(tmpPath);
    } else if (isPng && isLargeImage && !metadata.hasAlpha) {
      await pipeline.jpeg({ quality: JPEG_QUALITY, mozjpeg: true }).toFile(tmpPath);
    } else if (isPng) {
      await pipeline.png({ compressionLevel: PNG_COMPRESSION }).toFile(tmpPath);
    } else {
      await pipeline.png({ compressionLevel: PNG_COMPRESSION }).toFile(tmpPath);
    }

    const optimizedSize = fs.statSync(tmpPath).size;
    if (optimizedSize < originalSize) {
      fs.renameSync(tmpPath, filePath);
    } else {
      fs.unlinkSync(tmpPath);
    }
  } catch {
    // If sharp fails, keep original file untouched
    const tmpPath = filePath + '.opt';
    if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath);
  }
}

// Walk parsed data and optimize all image files on disk.
// Processes in batches of 5 to limit concurrent disk I/O.
const IMAGE_BATCH_SIZE = 5;

export async function optimizeImages(parsedData) {
  const items = [];

  function walk(elements) {
    for (const el of elements) {
      if (el.type === 'image' && el.imageFile) {
        items.push(el);
      } else if (el.type === 'group' && el.elements) {
        walk(el.elements);
      }
    }
  }

  for (const slide of parsedData.slides) {
    walk(slide.elements);
  }

  for (let i = 0; i < items.length; i += IMAGE_BATCH_SIZE) {
    const batch = items.slice(i, i + IMAGE_BATCH_SIZE);
    await Promise.allSettled(
      batch.map(el => optimizeImageFile(el.imageFile))
    );
  }

  return parsedData;
}
