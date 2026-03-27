import sharp from 'sharp';

// Max width for images in a Word document (letter page ~6.5" printable at 150 DPI)
const MAX_WIDTH_PX = 975;
const MAX_HEIGHT_PX = 1200;
const JPEG_QUALITY = 80;
const PNG_COMPRESSION = 9;

// Optimize an image buffer for embedding in a Word document.
// - Resizes if larger than page dimensions
// - Compresses JPEG/PNG
// - Converts large PNGs (photos) to JPEG for better compression
export async function optimizeImage(buffer) {
  if (!buffer || buffer.length === 0) return buffer;

  try {
    const image = sharp(buffer);
    const metadata = await image.metadata();

    if (!metadata.width || !metadata.height) return buffer;

    let pipeline = image;

    // Resize if larger than max dimensions
    if (metadata.width > MAX_WIDTH_PX || metadata.height > MAX_HEIGHT_PX) {
      pipeline = pipeline.resize(MAX_WIDTH_PX, MAX_HEIGHT_PX, {
        fit: 'inside',
        withoutEnlargement: true,
      });
    }

    // Choose output format:
    // - Small PNGs (icons, diagrams with transparency) stay PNG
    // - Large PNGs that are likely photos convert to JPEG
    // - JPEGs stay JPEG with quality compression
    const isJpeg = metadata.format === 'jpeg' || metadata.format === 'jpg';
    const isPng = metadata.format === 'png';
    const isLargeImage = buffer.length > 100 * 1024; // > 100KB

    let optimized;
    if (isJpeg) {
      optimized = await pipeline.jpeg({ quality: JPEG_QUALITY, mozjpeg: true }).toBuffer();
    } else if (isPng && isLargeImage && !metadata.hasAlpha) {
      // Large opaque PNG → convert to JPEG for much better compression
      optimized = await pipeline.jpeg({ quality: JPEG_QUALITY, mozjpeg: true }).toBuffer();
    } else if (isPng) {
      optimized = await pipeline.png({ compressionLevel: PNG_COMPRESSION }).toBuffer();
    } else {
      // Other formats (GIF, WebP, etc.) → convert to PNG
      optimized = await pipeline.png({ compressionLevel: PNG_COMPRESSION }).toBuffer();
    }

    // Only use optimized version if it's actually smaller
    return optimized.length < buffer.length ? optimized : buffer;
  } catch {
    // If sharp fails (corrupted image, unsupported format), return original
    return buffer;
  }
}

// Walk parsed data and optimize all image buffers in place.
// Processes in batches of 5 to limit concurrent memory usage from sharp.
// Uses allSettled so a single failed image doesn't break the batch.
const IMAGE_BATCH_SIZE = 5;

export async function optimizeImages(parsedData) {
  const items = [];

  function walk(elements) {
    for (const el of elements) {
      if (el.type === 'image' && el.imageBuffer) {
        items.push(el);
      } else if (el.type === 'group' && el.elements) {
        walk(el.elements);
      }
    }
  }

  for (const slide of parsedData.slides) {
    walk(slide.elements);
  }

  // Process in batches to avoid OOM with many large images
  for (let i = 0; i < items.length; i += IMAGE_BATCH_SIZE) {
    const batch = items.slice(i, i + IMAGE_BATCH_SIZE);
    const results = await Promise.allSettled(
      batch.map(el => optimizeImage(el.imageBuffer))
    );
    for (let j = 0; j < batch.length; j++) {
      if (results[j].status === 'fulfilled') {
        batch[j].imageBuffer = results[j].value;
      }
      // If rejected, keep original buffer
    }
  }

  return parsedData;
}
