import JSZip from 'jszip';
import fs from 'node:fs';
import path from 'node:path';
import { DOMParser } from '@xmldom/xmldom';
import {
  child, children, descendant, descendants,
  attr, nsAttr, textContent, allChildren, matchTag, NS,
} from './xmlHelper.js';

const domParser = new DOMParser();

function parseXml(xmlString) {
  return domParser.parseFromString(xmlString, 'application/xml');
}

// ── Theme parsing ──────────────────────────────────────────────

function parseTheme(themeXml) {
  const doc = parseXml(themeXml);
  const clrScheme = descendant(doc, 'a:clrScheme');
  const fontScheme = descendant(doc, 'a:fontScheme');

  const colors = {};
  if (clrScheme) {
    const colorNames = [
      'dk1', 'dk2', 'lt1', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'folHlink',
    ];
    for (const name of colorNames) {
      const el = child(clrScheme, `a:${name}`);
      if (el) {
        const srgb = child(el, 'a:srgbClr');
        const sys = child(el, 'a:sysClr');
        if (srgb) colors[name] = attr(srgb, 'val');
        else if (sys) colors[name] = attr(sys, 'lastClr') || attr(sys, 'val');
      }
    }
  }

  const fonts = { major: null, minor: null };
  if (fontScheme) {
    const major = descendant(fontScheme, 'a:majorFont');
    const minor = descendant(fontScheme, 'a:minorFont');
    if (major) {
      const latin = child(major, 'a:latin');
      if (latin) fonts.major = attr(latin, 'typeface');
    }
    if (minor) {
      const latin = child(minor, 'a:latin');
      if (latin) fonts.minor = attr(latin, 'typeface');
    }
  }

  return { colors, fonts };
}

// ── Color map parsing ──────────────────────────────────────────

function parseColorMap(node) {
  if (!node) return {};
  const map = {};
  const attrs = [
    'bg1', 'tx1', 'bg2', 'tx2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
  ];
  for (const a of attrs) {
    const val = attr(node, a);
    if (val) map[a] = val;
  }
  return map;
}

// ── Slide master style parsing ─────────────────────────────────

function parseMasterTextStyles(masterXml) {
  const doc = parseXml(masterXml);
  const txStyles = descendant(doc, 'p:txStyles');
  const styles = { title: {}, body: {}, other: {} };

  if (txStyles) {
    styles.title = extractLevelStyles(child(txStyles, 'p:titleStyle'));
    styles.body = extractLevelStyles(child(txStyles, 'p:bodyStyle'));
    styles.other = extractLevelStyles(child(txStyles, 'p:otherStyle'));
  }

  const clrMap = descendant(doc, 'p:clrMap');
  styles.colorMap = parseColorMap(clrMap);

  return styles;
}

function extractLevelStyles(styleNode) {
  if (!styleNode) return {};
  const levels = {};
  for (let i = 1; i <= 9; i++) {
    const lvl = child(styleNode, `a:lvl${i}pPr`);
    if (lvl) {
      const defRPr = child(lvl, 'a:defRPr');
      levels[i] = extractRunProps(defRPr);
      levels[i].alignment = attr(lvl, 'algn') || null;
    }
  }
  return levels;
}

// ── Slide layout parsing ───────────────────────────────────────

function parseLayoutStyles(layoutXml) {
  const doc = parseXml(layoutXml);
  const placeholders = {};

  const shapes = descendants(doc, 'p:sp');
  for (const sp of shapes) {
    const nvPr = descendant(sp, 'p:nvPr');
    const ph = nvPr ? child(nvPr, 'p:ph') : null;
    if (!ph) continue;

    const type = attr(ph, 'type') || 'body';
    const idx = attr(ph, 'idx');
    const txBody = child(sp, 'p:txBody');
    if (!txBody) continue;

    const lstStyle = child(txBody, 'a:lstStyle');
    const levels = {};
    if (lstStyle) {
      for (let i = 1; i <= 9; i++) {
        const lvl = child(lstStyle, `a:lvl${i}pPr`);
        if (lvl) {
          const defRPr = child(lvl, 'a:defRPr');
          levels[i] = extractRunProps(defRPr);
        }
      }
    }

    // Also extract default run props from paragraphs in the layout
    const paragraphs = children(txBody, 'a:p');
    const paraDefaults = [];
    for (const p of paragraphs) {
      const pPr = child(p, 'a:pPr');
      const endParaRPr = child(p, 'a:endParaRPr');
      const defRPr = pPr ? child(pPr, 'a:defRPr') : null;
      paraDefaults.push({
        level: pPr ? parseInt(attr(pPr, 'lvl') || '0', 10) : 0,
        props: extractRunProps(defRPr || endParaRPr),
      });
    }

    placeholders[type + (idx ? `_${idx}` : '')] = { levels, paraDefaults };
  }

  return placeholders;
}

// ── Relationship parsing ───────────────────────────────────────

function parseRels(relsXml) {
  const doc = parseXml(relsXml);
  const rels = {};
  const relEls = descendants(doc, 'rel:Relationship');
  // Fallback: some rels files use default namespace without prefix
  const all = relEls.length > 0 ? relEls : doc.getElementsByTagName('Relationship');
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    const id = el.getAttribute('Id');
    const target = el.getAttribute('Target');
    const type = el.getAttribute('Type');
    if (id && target) rels[id] = { target, type };
  }
  return rels;
}

// ── Run property extraction ────────────────────────────────────

function extractRunProps(rPr) {
  if (!rPr) return {};

  const props = {};

  const sz = attr(rPr, 'sz');
  if (sz) props.size = parseInt(sz, 10); // hundredths of a point

  const b = attr(rPr, 'b');
  if (b !== null) props.bold = b === '1' || b === 'true';

  const i = attr(rPr, 'i');
  if (i !== null) props.italic = i === '1' || i === 'true';

  const u = attr(rPr, 'u');
  if (u && u !== 'none') props.underline = u;

  const strike = attr(rPr, 'strike');
  if (strike && strike !== 'noStrike') props.strikethrough = strike;

  // Font
  const latin = child(rPr, 'a:latin');
  if (latin) {
    const typeface = attr(latin, 'typeface');
    if (typeface) props.font = typeface;
  }

  // Color
  const solidFill = child(rPr, 'a:solidFill');
  if (solidFill) {
    const color = extractColor(solidFill);
    if (color) props.color = color;
  }

  return props;
}

function extractColor(fillNode) {
  if (!fillNode) return null;

  const srgb = child(fillNode, 'a:srgbClr');
  if (srgb) {
    const val = attr(srgb, 'val');
    return val ? { type: 'rgb', value: val } : null;
  }

  const scheme = child(fillNode, 'a:schemeClr');
  if (scheme) {
    const val = attr(scheme, 'val');
    const transforms = extractColorTransforms(scheme);
    return val ? { type: 'scheme', value: val, transforms } : null;
  }

  const sys = child(fillNode, 'a:sysClr');
  if (sys) {
    const lastClr = attr(sys, 'lastClr');
    return lastClr ? { type: 'rgb', value: lastClr } : null;
  }

  return null;
}

function extractColorTransforms(colorNode) {
  const transforms = [];
  const transformNames = ['lumMod', 'lumOff', 'tint', 'shade', 'satMod', 'satOff', 'alpha'];
  for (const name of transformNames) {
    const el = child(colorNode, `a:${name}`);
    if (el) {
      const val = attr(el, 'val');
      if (val) transforms.push({ type: name, value: parseInt(val, 10) });
    }
  }
  return transforms;
}

// ── Resolve scheme color to RGB ────────────────────────────────

function resolveSchemeColor(schemeVal, colorMap, themeColors) {
  // Map through clrMap (e.g., tx1 → lt1)
  let resolved = colorMap[schemeVal] || schemeVal;
  return themeColors[resolved] || null;
}

function applyColorTransforms(hexColor, transforms) {
  if (!transforms || transforms.length === 0) return hexColor;

  let r = parseInt(hexColor.substring(0, 2), 16);
  let g = parseInt(hexColor.substring(2, 4), 16);
  let b = parseInt(hexColor.substring(4, 6), 16);

  for (const t of transforms) {
    const factor = t.value / 100000;
    switch (t.type) {
      case 'lumMod': {
        // Modulate luminance
        const { h, s, l } = rgbToHsl(r, g, b);
        const newL = Math.min(1, l * factor);
        ({ r, g, b } = hslToRgb(h, s, newL));
        break;
      }
      case 'lumOff': {
        const { h, s, l } = rgbToHsl(r, g, b);
        const newL = Math.min(1, Math.max(0, l + factor));
        ({ r, g, b } = hslToRgb(h, s, newL));
        break;
      }
      case 'tint': {
        r = Math.round(r + (255 - r) * factor);
        g = Math.round(g + (255 - g) * factor);
        b = Math.round(b + (255 - b) * factor);
        break;
      }
      case 'shade': {
        r = Math.round(r * factor);
        g = Math.round(g * factor);
        b = Math.round(b * factor);
        break;
      }
      // satMod, satOff, alpha: skip for now (minor visual impact)
    }
  }

  r = Math.max(0, Math.min(255, r));
  g = Math.max(0, Math.min(255, g));
  b = Math.max(0, Math.min(255, b));

  return r.toString(16).padStart(2, '0')
    + g.toString(16).padStart(2, '0')
    + b.toString(16).padStart(2, '0');
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function hslToRgb(h, s, l) {
  if (s === 0) {
    const v = Math.round(l * 255);
    return { r: v, g: v, b: v };
  }
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1/6) return p + (q - p) * 6 * t;
    if (t < 1/2) return q;
    if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: Math.round(hue2rgb(p, q, h + 1/3) * 255),
    g: Math.round(hue2rgb(p, q, h) * 255),
    b: Math.round(hue2rgb(p, q, h - 1/3) * 255),
  };
}

// ── Unsupported object detection ───────────────────────────────

const UNSUPPORTED_URIS = [
  { pattern: 'chart', label: 'Chart' },
  { pattern: 'diagram', label: 'SmartArt' },
  { pattern: 'oleObject', label: 'Embedded Object' },
  { pattern: 'wordprocessingDrawing', label: 'Embedded Drawing' },
];

function detectUnsupported(graphicFrame) {
  const cNvPr = descendant(graphicFrame, 'p:cNvPr');
  const name = cNvPr ? attr(cNvPr, 'name') : 'Unknown';

  const graphicData = descendant(graphicFrame, 'a:graphicData');
  if (!graphicData) return { type: 'unsupported', description: name || 'Unknown Object' };

  const uri = attr(graphicData, 'uri') || '';

  // Tables are in graphicFrame but we can attempt to extract text from them
  if (uri.includes('table')) return null; // handled separately

  for (const { pattern, label } of UNSUPPORTED_URIS) {
    if (uri.includes(pattern)) {
      return { type: 'unsupported', description: `${label}${name ? ': ' + name : ''}` };
    }
  }

  return { type: 'unsupported', description: name || 'Unknown Object' };
}

// ── Table extraction ───────────────────────────────────────────

function extractTable(graphicFrame) {
  const tbl = descendant(graphicFrame, 'a:tbl');
  if (!tbl) return null;

  const rows = [];
  for (const tr of children(tbl, 'a:tr')) {
    const cells = [];
    for (const tc of children(tr, 'a:tc')) {
      const txBody = child(tc, 'a:txBody');
      if (txBody) {
        cells.push(extractTextBody(txBody));
      }
    }
    rows.push(cells);
  }
  return { type: 'table', rows };
}

// ── Shape extraction ───────────────────────────────────────────

function extractTextBody(txBody) {
  if (!txBody) return [];

  const paragraphs = [];
  for (const p of children(txBody, 'a:p')) {
    const pPr = child(p, 'a:pPr');
    const level = pPr ? parseInt(attr(pPr, 'lvl') || '0', 10) : 0;
    const alignment = pPr ? attr(pPr, 'algn') : null;

    // Paragraph-level default run props
    const pDefRPr = pPr ? child(pPr, 'a:defRPr') : null;
    const paraDefaultProps = extractRunProps(pDefRPr);

    const runs = [];

    for (const c of allChildren(p)) {
      if (matchTag(c, 'a:r')) {
        const rPr = child(c, 'a:rPr');
        const t = child(c, 'a:t');
        const text = t ? textContent(t) : '';
        // Preserve space exactly — PPTX uses xml:space="preserve"
        runs.push({
          type: 'text',
          text,
          props: extractRunProps(rPr),
        });
      } else if (matchTag(c, 'a:br')) {
        runs.push({ type: 'break' });
      } else if (matchTag(c, 'a:fld')) {
        // Field (slide number, date, etc.) — extract text verbatim
        const t = child(c, 'a:t');
        const text = t ? textContent(t) : '';
        const rPr = child(c, 'a:rPr');
        if (text) {
          runs.push({
            type: 'text',
            text,
            props: extractRunProps(rPr),
          });
        }
      }
    }

    paragraphs.push({ level, alignment, defaultProps: paraDefaultProps, runs });
  }

  return paragraphs;
}

// Extract position and dimensions from spPr/xfrm (shared across shape types).
// graphicFrame uses p:xfrm directly instead of p:spPr/a:xfrm.
function extractPosition(shapeNode) {
  const spPr = child(shapeNode, 'p:spPr') || child(shapeNode, 'p:grpSpPr');
  const xfrm = (spPr ? child(spPr, 'a:xfrm') : null) || child(shapeNode, 'p:xfrm');
  const off = xfrm ? child(xfrm, 'a:off') : null;
  const ext = xfrm ? child(xfrm, 'a:ext') : null;

  return {
    x: off ? parseInt(attr(off, 'x') || '0', 10) : 0,
    y: off ? parseInt(attr(off, 'y') || '0', 10) : 0,
    cx: ext ? parseInt(attr(ext, 'cx') || '0', 10) : 0,
    cy: ext ? parseInt(attr(ext, 'cy') || '0', 10) : 0,
  };
}

function extractShape(sp) {
  const cNvPr = descendant(sp, 'p:cNvPr');
  const name = cNvPr ? attr(cNvPr, 'name') : '';
  const altText = cNvPr ? attr(cNvPr, 'descr') : null;

  // Placeholder info
  const nvPr = descendant(sp, 'p:nvPr');
  const ph = nvPr ? child(nvPr, 'p:ph') : null;
  const phType = ph ? (attr(ph, 'type') || 'body') : null;
  const phIdx = ph ? attr(ph, 'idx') : null;

  const txBody = child(sp, 'p:txBody');
  const paragraphs = extractTextBody(txBody);
  const position = extractPosition(sp);

  return {
    type: 'shape',
    name,
    altText,
    placeholder: phType ? { type: phType, idx: phIdx } : null,
    paragraphs,
    position,
  };
}

function extractPicture(pic, slideRels) {
  const cNvPr = descendant(pic, 'p:cNvPr');
  const altText = cNvPr ? attr(cNvPr, 'descr') : null;
  const name = cNvPr ? attr(cNvPr, 'name') : '';

  const blipFill = child(pic, 'p:blipFill');
  const blip = blipFill ? child(blipFill, 'a:blip') : null;
  const embedId = blip ? nsAttr(blip, 'r:embed') : null;

  let imagePath = null;
  if (embedId && slideRels[embedId]) {
    imagePath = slideRels[embedId].target;
  }

  const position = extractPosition(pic);

  return {
    type: 'image',
    name,
    altText: altText || null, // preserve existing, never add
    imagePath,
    embedId,
    width: position.cx,   // EMUs
    height: position.cy,  // EMUs
    position,
  };
}

// ── Recursive shape tree traversal ─────────────────────────────

function extractElements(spTree, slideRels) {
  const elements = [];

  for (const node of allChildren(spTree)) {
    if (matchTag(node, 'p:sp')) {
      elements.push(extractShape(node));
    } else if (matchTag(node, 'p:pic')) {
      elements.push(extractPicture(node, slideRels));
    } else if (matchTag(node, 'p:grpSp')) {
      // Recurse into group shapes
      const groupElements = extractElements(node, slideRels);
      const position = extractPosition(node);
      elements.push({ type: 'group', elements: groupElements, position });
    } else if (matchTag(node, 'p:graphicFrame')) {
      elements.push(...extractGraphicFrame(node, slideRels));
    } else if (matchTag(node, 'mc:AlternateContent')) {
      // mc:AlternateContent wraps SmartArt/charts with a fallback image.
      // Try mc:Fallback first (contains a p:pic with the cached rendering),
      // then fall back to mc:Choice (contains the raw graphicFrame).
      const fallback = child(node, 'mc:Fallback');
      const choice = child(node, 'mc:Choice');

      if (fallback) {
        // Look for a p:pic (cached image) in the fallback
        const pic = child(fallback, 'p:pic');
        if (pic) {
          const img = extractPicture(pic, slideRels);
          img.isSmartArtFallback = true;
          elements.push(img);
          continue;
        }
        // Or recurse into fallback for other shape types
        elements.push(...extractElements(fallback, slideRels));
      } else if (choice) {
        // No fallback — try the choice content
        elements.push(...extractElements(choice, slideRels));
      }
    }
  }

  return elements;
}

function extractGraphicFrame(node, slideRels) {
  const position = extractPosition(node);
  // Check for table first
  const table = extractTable(node);
  if (table) {
    table.position = position;
    return [table];
  }

  // Check if this is SmartArt — try to extract as image from drawing cache
  const graphicData = descendant(node, 'a:graphicData');
  const uri = graphicData ? attr(graphicData, 'uri') || '' : '';

  if (uri.includes('diagram')) {
    // SmartArt: look for the drawing relationship (r:dm points to data,
    // but we need the drawing cache). Check slide rels for a diagramDrawing
    // relationship tied to this frame's relationships.
    const dgmRelIds = descendant(node, 'dgm:relIds');
    if (dgmRelIds) {
      // Try to find a drawing cache image via the data model relationship
      const dmId = nsAttr(dgmRelIds, 'r:dm');
      if (dmId && slideRels[dmId]) {
        // Store diagram info so extractImages can find the cached rendering
        return [{
          type: 'smartart',
          description: `SmartArt`,
          position,
          diagramRels: {
            dm: dmId,
            lo: nsAttr(dgmRelIds, 'r:lo'),
            qs: nsAttr(dgmRelIds, 'r:qs'),
            cs: nsAttr(dgmRelIds, 'r:cs'),
          },
          slideRels,
          width: position.cx,
          height: position.cy,
        }];
      }
    }
  }

  const unsupported = detectUnsupported(node);
  if (unsupported) {
    unsupported.position = position;
    return [unsupported];
  }
  return [];
}

// ── Main parse function ────────────────────────────────────────

const MAX_UNCOMPRESSED_SIZE = 200 * 1024 * 1024; // 200MB limit on decompressed content

export async function parsePptx(buffer) {
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (err) {
    throw new Error('Invalid or corrupted PPTX file: ' + err.message);
  }

  // Guard against zip bombs: check total uncompressed size
  let totalUncompressed = 0;
  zip.forEach((_path, file) => {
    if (!file.dir && file._data?.uncompressedSize) {
      totalUncompressed += file._data.uncompressedSize;
    }
  });
  if (totalUncompressed > MAX_UNCOMPRESSED_SIZE) {
    throw new Error('PPTX file contains too much data (possible zip bomb)');
  }

  // 1. Parse presentation.xml to get slide order
  const presentationFile = zip.file('ppt/presentation.xml');
  if (!presentationFile) {
    throw new Error('Invalid PPTX: missing ppt/presentation.xml');
  }
  const presentationXml = await presentationFile.async('text');
  const presDoc = parseXml(presentationXml);
  const sldIdLst = descendant(presDoc, 'p:sldIdLst');
  const sldIds = sldIdLst ? children(sldIdLst, 'p:sldId') : [];

  // 2. Parse presentation relationships
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
  if (!presRelsFile) {
    throw new Error('Invalid PPTX: missing presentation relationships');
  }
  const presRelsXml = await presRelsFile.async('text');
  const presRels = parseRels(presRelsXml);

  // 3. Parse theme
  let theme = { colors: {}, fonts: { major: null, minor: null } };
  const themeFile = zip.file('ppt/theme/theme1.xml');
  if (themeFile) {
    const themeXml = await themeFile.async('text');
    theme = parseTheme(themeXml);
  }

  // 4. Parse slide master(s) for text styles and color map
  let masterStyles = { title: {}, body: {}, other: {}, colorMap: {} };
  const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
  if (masterFile) {
    const masterXml = await masterFile.async('text');
    masterStyles = parseMasterTextStyles(masterXml);
  }

  // 5. Parse each slide in order
  const slides = [];
  const unsupportedObjects = [];

  for (let i = 0; i < sldIds.length; i++) {
    const sldId = sldIds[i];
    const rId = nsAttr(sldId, 'r:id') || attr(sldId, 'r:id');
    if (!rId || !presRels[rId]) continue;

    const slidePath = 'ppt/' + presRels[rId].target.replace(/^\.?\//, '');
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;

    const slideXml = await slideFile.async('text');
    const slideDoc = parseXml(slideXml);

    // Parse slide relationships
    const slideNum = slidePath.match(/slide(\d+)\.xml/)?.[1];
    const slideRelsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const slideRelsFile = zip.file(slideRelsPath);
    const slideRels = slideRelsFile ? parseRels(await slideRelsFile.async('text')) : {};

    // Get layout reference for style inheritance
    let layoutStyles = {};
    const layoutRel = Object.values(slideRels).find(r => r.type?.includes('slideLayout'));
    if (layoutRel) {
      const layoutPath = 'ppt/slideLayouts/' + layoutRel.target.replace(/^.*\//, '');
      const layoutFile = zip.file(layoutPath);
      if (layoutFile) {
        const layoutXml = await layoutFile.async('text');
        layoutStyles = parseLayoutStyles(layoutXml);
      }
    }

    // Check for color map override
    const clrMapOvr = descendant(slideDoc, 'p:clrMapOvr');
    let slideColorMap = masterStyles.colorMap;
    if (clrMapOvr) {
      const override = child(clrMapOvr, 'a:overrideClrMapping');
      if (override) {
        slideColorMap = { ...masterStyles.colorMap, ...parseColorMap(override) };
      }
    }

    // Extract shape tree
    const cSld = descendant(slideDoc, 'p:cSld');
    const spTree = cSld ? child(cSld, 'p:spTree') : null;
    const elements = spTree ? extractElements(spTree, slideRels) : [];

    // Collect unsupported objects
    for (const el of flattenElements(elements)) {
      if (el.type === 'unsupported') {
        unsupportedObjects.push({ slide: i + 1, description: el.description });
      }
    }

    slides.push({
      index: i + 1,
      elements,
      colorMap: slideColorMap,
      layoutStyles,
    });
  }

  return {
    slides,
    theme,
    masterStyles,
    unsupportedObjects,
    zip, // reuse in extractImages to avoid re-loading buffer
  };
}

// Flatten nested element structures (groups) into a flat list
function flattenElements(elements) {
  const flat = [];
  for (const el of elements) {
    flat.push(el);
    if (el.type === 'group' && el.elements) {
      flat.push(...flattenElements(el.elements));
    }
    if (el.type === 'table' && el.rows) {
      for (const row of el.rows) {
        for (const cell of row) {
          // cells are arrays of paragraphs, not elements
        }
      }
    }
  }
  return flat;
}

// ── Resolve all colors to RGB ──────────────────────────────────

export function resolveColors(parsedData) {
  const { slides, theme, masterStyles } = parsedData;

  for (const slide of slides) {
    const colorMap = slide.colorMap || masterStyles.colorMap || {};
    resolveElementColors(slide.elements, colorMap, theme.colors);
  }

  return parsedData;
}

function resolveElementColors(elements, colorMap, themeColors) {
  for (const el of elements) {
    if (el.type === 'shape' && el.paragraphs) {
      for (const para of el.paragraphs) {
        resolveRunColors(para.runs, colorMap, themeColors);
        resolvePropsColor(para.defaultProps, colorMap, themeColors);
      }
    } else if (el.type === 'group' && el.elements) {
      resolveElementColors(el.elements, colorMap, themeColors);
    } else if (el.type === 'table' && el.rows) {
      for (const row of el.rows) {
        for (const cell of row) {
          for (const para of cell) {
            resolveRunColors(para.runs, colorMap, themeColors);
            resolvePropsColor(para.defaultProps, colorMap, themeColors);
          }
        }
      }
    }
  }
}

function resolveRunColors(runs, colorMap, themeColors) {
  for (const run of runs) {
    if (run.props) {
      resolvePropsColor(run.props, colorMap, themeColors);
    }
  }
}

function resolvePropsColor(props, colorMap, themeColors) {
  if (!props || !props.color) return;
  if (props.color.type === 'scheme') {
    const rgb = resolveSchemeColor(props.color.value, colorMap, themeColors);
    if (rgb) {
      const finalRgb = applyColorTransforms(rgb, props.color.transforms);
      props.color = { type: 'rgb', value: finalRgb };
    }
  }
}

// ── Style inheritance resolution ───────────────────────────────

export function resolveInheritedStyles(parsedData) {
  const { slides, masterStyles, theme } = parsedData;

  for (const slide of slides) {
    for (const el of flattenElements(slide.elements)) {
      if (el.type !== 'shape' || !el.paragraphs) continue;

      // Determine which master style set applies
      const phType = el.placeholder?.type;
      let masterLevels = masterStyles.body;
      if (phType === 'title' || phType === 'ctrTitle') {
        masterLevels = masterStyles.title;
      } else if (!phType) {
        masterLevels = masterStyles.other;
      }

      // Get layout-level styles if this is a placeholder
      const layoutKey = phType ? phType + (el.placeholder?.idx ? `_${el.placeholder.idx}` : '') : null;
      const layoutPh = layoutKey ? slide.layoutStyles[layoutKey] : null;

      for (const para of el.paragraphs) {
        const level = (para.level || 0) + 1; // levels are 1-indexed in master styles

        // Build inherited defaults: theme → master → layout → paragraph default
        const masterDefault = masterLevels[level] || masterLevels[1] || {};
        const layoutDefault = layoutPh?.levels[level] || {};

        const inherited = {
          font: theme.fonts.minor,
          ...masterDefault,
          ...stripUndefined(layoutDefault),
          ...stripUndefined(para.defaultProps),
        };

        // Apply to runs that lack properties
        for (const run of para.runs) {
          if (run.type !== 'text') continue;
          run.props = {
            ...inherited,
            ...stripUndefined(run.props),
          };
        }
      }
    }
  }

  return parsedData;
}

function stripUndefined(obj) {
  if (!obj) return {};
  const result = {};
  for (const [key, val] of Object.entries(obj)) {
    if (val !== undefined && val !== null) result[key] = val;
  }
  return result;
}

// ── Extract image buffers ──────────────────────────────────────

// Safely resolve a relative image path within the PPTX zip.
// Prevents path traversal attacks by ensuring the resolved path stays under ppt/.
function safeZipPath(imagePath) {
  // Reject absolute paths
  if (imagePath.startsWith('/')) return null;

  // Resolve relative segments (../media/image1.png → ppt/media/image1.png)
  const segments = ('ppt/slides/' + imagePath).split('/');
  const resolved = [];
  for (const seg of segments) {
    if (seg === '..') {
      if (resolved.length > 0) resolved.pop();
    } else if (seg !== '.' && seg !== '') {
      resolved.push(seg);
    }
  }
  const result = resolved.join('/');
  // Images must resolve to ppt/media/ (the standard location in PPTX)
  if (!result.startsWith('ppt/media/')) return null;
  return result;
}

// Extract images from the PPTX zip to disk files in workDir.
// If workDir is provided, writes files to disk (el.imageFile = path).
// If workDir is null, falls back to in-memory buffers (el.imageBuffer) for tests.
export async function extractImages(zipOrBuffer, parsedData, workDir) {
  const zip = zipOrBuffer instanceof JSZip ? zipOrBuffer : await JSZip.loadAsync(zipOrBuffer);
  const extracted = {}; // zipPath → diskPath or buffer

  for (const slide of parsedData.slides) {
    for (const el of flattenElements(slide.elements)) {
      if (el.type === 'image' && el.imagePath) {
        const normalized = safeZipPath(el.imagePath);
        if (!normalized) {
          el.imageFile = null;
          el.imageBuffer = null;
          continue;
        }

        if (!extracted[normalized]) {
          const file = zip.file(normalized);
          if (file) {
            const buf = await file.async('nodebuffer');
            if (workDir) {
              const filename = path.basename(normalized);
              const diskPath = path.join(workDir, filename);
              await fs.promises.writeFile(diskPath, buf);
              extracted[normalized] = { file: diskPath };
            } else {
              extracted[normalized] = { buffer: buf };
            }
          }
        }

        const entry = extracted[normalized];
        if (entry) {
          el.imageFile = entry.file || null;
          el.imageBuffer = entry.buffer || null;
        } else {
          el.imageFile = null;
          el.imageBuffer = null;
        }
        el.resolvedPath = normalized;
      }
    }
  }

  return parsedData;
}

