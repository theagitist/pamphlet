// Maps parsed PPTX run properties to docx package TextRun options.

import { UnderlineType, AlignmentType } from 'docx';

// PPTX sz is in hundredths of a point (e.g., 2400 = 24pt).
// docx TextRun size is in half-points (e.g., 48 = 24pt).
export function sizeToHalfPoints(sz) {
  if (!sz) return undefined;
  return Math.round(sz / 50);
}

// PPTX underline values → docx UnderlineType
const UNDERLINE_MAP = {
  sng: UnderlineType.SINGLE,
  dbl: UnderlineType.DOUBLE,
  heavy: UnderlineType.THICK,
  dotted: UnderlineType.DOTTED,
  dottedHeavy: UnderlineType.DOTTED_HEAVY,
  dash: UnderlineType.DASH,
  dashHeavy: UnderlineType.DASHED_HEAVY,
  dashLong: UnderlineType.DASH_LONG,
  dashLongHeavy: UnderlineType.DASH_LONG_HEAVY,
  dotDash: UnderlineType.DOT_DASH,
  dotDashHeavy: UnderlineType.DOT_DOT_DASH_HEAVY,
  wavy: UnderlineType.WAVE,
  wavyHeavy: UnderlineType.WAVY_HEAVY,
  wavyDbl: UnderlineType.WAVY_DOUBLE,
};

export function mapUnderline(pptxUnderline) {
  if (!pptxUnderline) return undefined;
  return UNDERLINE_MAP[pptxUnderline] || UnderlineType.SINGLE;
}

// PPTX alignment → docx AlignmentType
const ALIGN_MAP = {
  l: AlignmentType.LEFT,
  r: AlignmentType.RIGHT,
  ctr: AlignmentType.CENTER,
  just: AlignmentType.JUSTIFIED,
  dist: AlignmentType.DISTRIBUTE,
};

export function mapAlignment(pptxAlign) {
  if (!pptxAlign) return undefined;
  return ALIGN_MAP[pptxAlign] || undefined;
}

// ── Typography Scaler ──────────────────────────────────────────
//
// PPTX slides use large fonts for projection (40-50pt titles, 18-24pt body).
// Word documents need smaller type. The scaler:
//   1. Finds the max font size across all runs in the deck
//   2. Maps that max to TARGET_MAX_PT (22pt for titles)
//   3. Scales everything proportionally
//   4. Clamps: floor at MIN_PT (10pt), body cap at MAX_BODY_PT (14pt)
//
// "Body" is detected as any non-bold run or any run below the 60th percentile
// of the original size range. Bold/large runs are treated as headers and get
// the upper range of the scale.

const TARGET_MAX_PT = 22;   // largest heading in output
const MAX_BODY_PT = 14;     // body text cap
const MIN_PT = 10;          // nothing smaller than this

export function buildTypographyScaler(parsedData) {
  // Collect all sizes (in hundredths of a point) from the deck
  const allSizes = [];
  function collect(elements) {
    for (const el of elements) {
      if (el.type === 'shape' && el.paragraphs) {
        for (const para of el.paragraphs) {
          for (const run of para.runs) {
            if (run.type === 'text' && run.props?.size) allSizes.push(run.props.size);
          }
        }
      } else if (el.type === 'group' && el.elements) {
        collect(el.elements);
      } else if (el.type === 'table' && el.rows) {
        for (const row of el.rows) {
          for (const cell of row) {
            for (const para of cell) {
              for (const run of para.runs) {
                if (run.type === 'text' && run.props?.size) allSizes.push(run.props.size);
              }
            }
          }
        }
      }
    }
  }
  for (const slide of parsedData.slides) collect(slide.elements);

  if (allSizes.length === 0) return (sz) => sz;

  const maxSz = Math.max(...allSizes);
  const minSz = Math.min(...allSizes);

  // Sort to find body threshold (60th percentile)
  const sorted = [...allSizes].sort((a, b) => a - b);
  const bodyThresholdSz = sorted[Math.floor(sorted.length * 0.6)];

  const maxPt = maxSz / 100;
  const bodyThresholdPt = bodyThresholdSz / 100;

  // Return a function that scales a size (in hundredths of pt) → hundredths of pt
  return function scaleSize(sz) {
    if (!sz) return sz;
    const pt = sz / 100;

    // Proportional scale: original pt → scaled pt
    let scaled;
    if (maxPt === minSz / 100) {
      // All same size — just use body cap
      scaled = MAX_BODY_PT;
    } else {
      // Linear scale from [minPt..maxPt] → [MIN_PT..TARGET_MAX_PT]
      const minPt = minSz / 100;
      const ratio = (pt - minPt) / (maxPt - minPt);
      scaled = MIN_PT + ratio * (TARGET_MAX_PT - MIN_PT);
    }

    // Cap body-sized text
    if (pt <= bodyThresholdPt) {
      scaled = Math.min(scaled, MAX_BODY_PT);
    }

    // Enforce floor and ceiling
    scaled = Math.max(MIN_PT, Math.min(TARGET_MAX_PT, scaled));

    // Round to nearest 0.5pt
    scaled = Math.round(scaled * 2) / 2;

    return scaled * 100; // back to hundredths of pt
  };
}

export function applyTypographyScaling(parsedData) {
  const scale = buildTypographyScaler(parsedData);

  function scaleElements(elements) {
    for (const el of elements) {
      if (el.type === 'shape' && el.paragraphs) {
        for (const para of el.paragraphs) {
          for (const run of para.runs) {
            if (run.type === 'text' && run.props?.size) {
              run.props.size = scale(run.props.size);
            }
          }
          if (para.defaultProps?.size) {
            para.defaultProps.size = scale(para.defaultProps.size);
          }
        }
      } else if (el.type === 'group' && el.elements) {
        scaleElements(el.elements);
      } else if (el.type === 'table' && el.rows) {
        for (const row of el.rows) {
          for (const cell of row) {
            for (const para of cell) {
              for (const run of para.runs) {
                if (run.type === 'text' && run.props?.size) {
                  run.props.size = scale(run.props.size);
                }
              }
            }
          }
        }
      }
    }
  }

  for (const slide of parsedData.slides) {
    scaleElements(slide.elements);
  }

  return parsedData;
}

// Build docx TextRun options from parsed props
export function mapRunProps(props) {
  if (!props) return {};

  const options = {};

  if (props.font) options.font = props.font;
  if (props.size) options.size = sizeToHalfPoints(props.size);
  if (props.bold) options.bold = true;
  if (props.italic) options.italics = true; // docx uses 'italics' not 'italic'
  if (props.underline) {
    options.underline = { type: mapUnderline(props.underline) };
  }
  if (props.strikethrough) options.strike = true;

  if (props.color && props.color.type === 'rgb' && props.color.value) {
    options.color = props.color.value;
  }

  return options;
}
