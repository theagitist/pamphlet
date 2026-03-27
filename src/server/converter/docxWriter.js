import fs from 'node:fs';
import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle,
  Footer, PageNumber,
} from 'docx';
import { mapRunProps, mapAlignment } from './styleMapper.js';

// EMU to inches, then to docx points (1 inch = 72pt, but ImageRun uses pixels at 96dpi)
// docx ImageRun expects dimensions in pixels (at 96 DPI) or EMU.
// We'll pass EMU directly using the transformation object.
function emuToPixels(emu) {
  // 1 inch = 914400 EMU, 1 inch = 96 pixels (screen)
  return Math.round((emu / 914400) * 96);
}

// Standard slide dimensions in EMU (widescreen 13.33" x 7.5")
const SLIDE_WIDTH_EMU = 12192000;
const SLIDE_HEIGHT_EMU = 6858000;

// Filter out images that shouldn't appear in a Word handout:
// - Full-slide backgrounds (decorative, nearly full-slide size)
// - Tiny decorative shapes (< 4KB, usually icons/bullets from themes)
function shouldSkipImage(el) {
  // Skip full-slide-sized images (backgrounds)
  if (el.width >= SLIDE_WIDTH_EMU * 0.9 && el.height >= SLIDE_HEIGHT_EMU * 0.9) {
    return true;
  }

  // Skip tiny images (decorative shapes, bullet icons, etc.)
  const imageFile = el.imageFile || el.imageBuffer;
  if (imageFile) {
    const size = typeof imageFile === 'string'
      ? (fs.existsSync(imageFile) ? fs.statSync(imageFile).size : 0)
      : imageFile.length;
    if (size < 4096) return true;
  }

  return false;
}

function buildTextRun(run) {
  const opts = mapRunProps(run.props);
  opts.text = run.text;
  return new TextRun(opts);
}

// Detect if a paragraph looks like a header (bold text with larger font)
function isHeaderParagraph(para) {
  const textRuns = para.runs.filter(r => r.type === 'text' && r.text.trim());
  if (textRuns.length === 0) return false;

  // Check if the majority of text runs are bold
  const boldRuns = textRuns.filter(r => r.props?.bold);
  if (boldRuns.length < textRuns.length * 0.5) return false;

  // Check if font size is larger than typical body (> 14pt = 1400 hundredths)
  const avgSize = textRuns.reduce((sum, r) => sum + (r.props?.size || 0), 0) / textRuns.length;
  if (avgSize > 1400) return true;

  // Bold text at any size that has short content (< 80 chars) is likely a header
  const totalText = textRuns.map(r => r.text).join('').trim();
  if (boldRuns.length === textRuns.length && totalText.length < 80) return true;

  return false;
}

function buildParagraph(para, { keepNext = false } = {}) {
  const children = [];

  for (const run of para.runs) {
    if (run.type === 'text') {
      children.push(buildTextRun(run));
    } else if (run.type === 'break') {
      children.push(new TextRun({ break: 1 }));
    }
  }

  // If paragraph has no runs, add an empty run to preserve the empty line
  if (children.length === 0) {
    const opts = mapRunProps(para.defaultProps || {});
    opts.text = '';
    children.push(new TextRun(opts));
  }

  const paragraphOpts = { children };

  const alignment = mapAlignment(para.alignment);
  if (alignment) paragraphOpts.alignment = alignment;

  if (keepNext) {
    paragraphOpts.keepNext = true;
  }

  if (isHeaderParagraph(para)) {
    paragraphOpts.spacing = { before: 50, after: 240 };
  } else {
    // Body paragraphs: 160 twips (8pt) between them for readability
    paragraphOpts.spacing = { after: 160 };
  }

  return new Paragraph(paragraphOpts);
}

function buildUnsupportedPlaceholder(description, { keepNext = false } = {}) {
  return new Paragraph({
    children: [
      new TextRun({
        text: `[Object not supported: ${description}]`,
        italics: true,
        color: '999999',
        size: 20, // 10pt in half-points
      }),
    ],
    keepNext,
  });
}

function buildImageParagraph(el, { keepNext = false } = {}) {
  // Read image from disk file (not held in memory)
  const imageFile = el.imageFile || el.imageBuffer; // backward compat for tests
  if (!imageFile) {
    return buildUnsupportedPlaceholder(`Image: ${el.name || 'missing'}`, { keepNext });
  }

  // Read the image buffer from disk at assembly time, then let GC reclaim it
  const data = typeof imageFile === 'string' ? fs.readFileSync(imageFile) : imageFile;

  const width = el.width ? emuToPixels(el.width) : 400;
  const height = el.height ? emuToPixels(el.height) : 300;

  // Cap to reasonable page dimensions for a letter-size Word document.
  // Max width ~6.25" (600px at 96 DPI), max height ~6" (576px) so tall
  // images share the page with a header rather than consuming an entire page.
  const maxWidth = 600;
  const maxHeight = 576;
  let finalWidth = width;
  let finalHeight = height;
  if (finalWidth > maxWidth) {
    const scale = maxWidth / finalWidth;
    finalWidth = maxWidth;
    finalHeight = Math.round(finalHeight * scale);
  }
  if (finalHeight > maxHeight) {
    const scale = maxHeight / finalHeight;
    finalHeight = maxHeight;
    finalWidth = Math.round(finalWidth * scale);
  }

  const imageOpts = {
    type: 'png',
    data,
    transformation: { width: finalWidth, height: finalHeight },
  };

  if (el.altText) {
    imageOpts.altText = {
      title: el.altText,
      description: el.altText,
      name: el.name || 'Image',
    };
  }

  return new Paragraph({
    children: [new ImageRun(imageOpts)],
    keepNext,
  });
}

function buildTableParagraphs(tableEl) {
  const paragraphs = [];

  const rows = tableEl.rows.map(row => {
    const cells = row.map(cellParas => {
      const cellChildren = [];
      for (const para of cellParas) {
        cellChildren.push(buildParagraph(para, false));
      }
      if (cellChildren.length === 0) {
        cellChildren.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
      }
      return new TableCell({
        children: cellChildren,
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        },
      });
    });
    return new TableRow({ children: cells });
  });

  if (rows.length > 0) {
    paragraphs.push(new Table({
      rows,
      width: { size: 100, type: WidthType.PERCENTAGE },
    }));
  }

  return paragraphs;
}

// Sort elements by their spatial position on the slide (top-to-bottom, then left-to-right).
// This ensures the Word document reads in visual order regardless of how the PPTX
// source ordered shapes in its XML tree.
//
// Special handling:
// - y=0 on non-shape elements (tables, unsupported) is treated as a default/unknown
//   position and sorted after positioned elements at the same level
// - Elements within the same horizontal band (y within ROW_THRESHOLD EMU) are grouped
//   and sorted left-to-right, then their children read top-to-bottom within the column
function reorderElements(elements) {
  // Separate elements into three buckets:
  // - positioned: have real spatial coordinates, sorted by position
  // - unpositionedContent: tables/unsupported at y=0, inserted after title band
  // - trailingImages: full-height/side-panel images, appended at the end
  const positioned = [];
  const unpositionedContent = [];
  const trailingImages = [];

  for (const el of elements) {
    const y = el.position?.y ?? 0;
    const isFullHeightImage = el.type === 'image' && y === 0
      && el.height >= SLIDE_HEIGHT_EMU * 0.9;
    const isDefaultPos = y === 0 && el.type !== 'shape' && el.type !== 'image';
    if (isFullHeightImage) {
      trailingImages.push(el);
    } else if (isDefaultPos) {
      unpositionedContent.push(el);
    } else {
      positioned.push(el);
    }
  }

  // Sort positioned elements: top-to-bottom, left-to-right for same row
  // Elements within ~5% of slide height (~343000 EMU) of each other are "same row"
  const ROW_THRESHOLD = 350000;

  positioned.sort((a, b) => {
    const ay = a.position?.y ?? 0;
    const by = b.position?.y ?? 0;

    // If elements are in the same horizontal band, sort by x
    if (Math.abs(ay - by) < ROW_THRESHOLD) {
      const ax = a.position?.x ?? 0;
      const bx = b.position?.x ?? 0;
      return ax - bx;
    }

    return ay - by;
  });

  // Insert unpositioned content (tables, SmartArt) after the title band
  if (unpositionedContent.length > 0 && positioned.length > 0) {
    const firstTextEl = positioned.find(el =>
      el.type === 'shape' && el.paragraphs?.some(p =>
        p.runs.some(r => r.type === 'text' && r.text.trim())
      )
    );
    const titleY = firstTextEl?.position?.y ?? positioned[0].position?.y ?? 0;

    let insertIdx = positioned.findIndex(el =>
      (el.position?.y ?? 0) > titleY + ROW_THRESHOLD
    );
    if (insertIdx === -1) insertIdx = positioned.length;
    positioned.splice(insertIdx, 0, ...unpositionedContent);
    return [...positioned, ...trailingImages];
  }

  return [...positioned, ...unpositionedContent, ...trailingImages];
}

function buildElementParagraphs(elements, { keepNext = false } = {}) {
  const reordered = reorderElements(elements);
  const paragraphs = [];

  for (const el of reordered) {
    if (el.type === 'shape' && el.paragraphs) {
      // Skip shapes with no text content (empty decorative/grid shapes)
      const hasText = el.paragraphs.some(p =>
        p.runs.some(r => r.type === 'text' && r.text.trim())
      );
      if (!hasText) continue;

      for (let i = 0; i < el.paragraphs.length; i++) {
        paragraphs.push(buildParagraph(el.paragraphs[i], { keepNext }));
      }
    } else if (el.type === 'image') {
      if (shouldSkipImage(el)) continue;
      paragraphs.push(buildImageParagraph(el, { keepNext }));
    } else if (el.type === 'group' && el.elements) {
      paragraphs.push(...buildElementParagraphs(el.elements, { keepNext }));
    } else if (el.type === 'table') {
      paragraphs.push(...buildTableParagraphs(el));
    } else if (el.type === 'unsupported') {
      paragraphs.push(buildUnsupportedPlaceholder(el.description, { keepNext }));
    }
  }

  return paragraphs;
}

// Spacing between slides (in twentieths of a point). 360 twips = 18pt ≈ ~0.25 inch.
const SLIDE_SEPARATOR_SPACING = 180;

export async function writeDocx(parsedData, options = {}) {
  const { slides } = parsedData;

  // Build all slides into a single section with spacing between them
  // instead of forcing page breaks. Word handles pagination naturally.
  const allChildren = [];

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    // Build all paragraphs with keepNext so Word treats the slide as a unit.
    // If the slide doesn't fit in the remaining page space, Word starts a new page.
    const slideParagraphs = buildElementParagraphs(slide.elements, { keepNext: true });

    if (slideParagraphs.length === 0) continue;

    if (i > 0 && allChildren.length > 0) {
      // Spacer for vertical breathing room — no keepNext, so Word can break here.
      allChildren.push(new Paragraph({
        children: [],
        spacing: { before: SLIDE_SEPARATOR_SPACING },
      }));

      // Horizontal rule attached to the slide content via keepNext,
      // so the line always stays with the header below it.
      allChildren.push(new Paragraph({
        children: [],
        thematicBreak: true,
        keepNext: true,
      }));
    }

    // Add all slide paragraphs. Then append a terminator paragraph without
    // keepNext to break the chain, so the next slide's spacer is a clean break point.
    allChildren.push(...slideParagraphs);
    allChildren.push(new Paragraph({ children: [] }));
  }

  if (allChildren.length === 0) {
    allChildren.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
  }

  const sectionOpts = { children: allChildren };

  // Page numbers in footer (optional, controlled by options.pageNumbers)
  if (options.pageNumbers !== false) {
    sectionOpts.footers = {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ color: '999999', size: 18, children: [PageNumber.CURRENT] }),
            ],
          }),
        ],
      }),
    };
  }

  const doc = new Document({
    sections: [sectionOpts],
    features: { updateFields: true },
  });
  return Packer.toBuffer(doc);
}
