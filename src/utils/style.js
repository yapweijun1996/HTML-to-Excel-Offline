import { pxToExcelColWidth } from './measurement.js';

const BORDER_STYLE_MAP = {
  none: undefined,
  hidden: undefined,
  solid: 'thin',
  dotted: 'dotted',
  dashed: 'dashed',
  double: 'double',
  groove: 'thin',
  ridge: 'thin',
  inset: 'thin',
  outset: 'thin'
};

export function cssColorToARGB(color, fallback = 'FF000000') {
  if (!color || color === 'transparent') return fallback;
  const ctx = document.createElement('canvas').getContext('2d');
  if (!ctx) return fallback;
  ctx.fillStyle = '#000';
  ctx.fillStyle = color;
  const computed = ctx.fillStyle;
  const match = /^#([0-9a-f]{6})$/i.exec(computed);
  if (match) {
    return `FF${match[1].toUpperCase()}`;
  }
  if (/^#([0-9a-f]{3})$/i.test(computed)) {
    const hex = computed.slice(1).split('').map(ch => ch + ch).join('');
    return `FF${hex.toUpperCase()}`;
  }
  return fallback;
}

export function extractCellStyle(el, { header = false } = {}) {
  if (!el) {
    return {
      alignment: header ? { horizontal: 'center', vertical: 'middle' } : { horizontal: 'left', vertical: 'top' },
      font: header ? { bold: true } : undefined,
      border: defaultBorder(),
      fill: header ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } } : undefined
    };
  }
  const cs = getComputedStyle(el);
  const font = buildFont(cs, header);
  const alignment = buildAlignment(cs, header);
  const border = buildBorder(cs);
  const fill = buildFill(cs, header);
  return { alignment, font, border, fill };
}

export function buildColumnWidthFromElement(el) {
  if (!el) return { width: 16 };
  const rect = el.getBoundingClientRect();
  const widthPx = rect.width || parseCssWidth(el) || 120;
  return { width: pxToExcelColWidth(widthPx) };
}

function buildFont(cs, header) {
  const font = {};
  const family = cs.fontFamily && cs.fontFamily.split(',')[0].replace(/['"]/g, '').trim();
  if (family) font.name = family;
  const size = parseFloat(cs.fontSize);
  if (!Number.isNaN(size)) font.size = Math.round(size * 100) / 100;
  if (header || cs.fontWeight === 'bold' || parseInt(cs.fontWeight, 10) >= 600) font.bold = true;
  if (cs.fontStyle === 'italic') font.italic = true;
  if (cs.textDecorationLine && cs.textDecorationLine.includes('underline')) font.underline = true;
  if (Object.keys(font).length === 0) return undefined;
  return font;
}

function buildAlignment(cs, header) {
  const horizontal = cs.textAlign || (header ? 'center' : 'left');
  const vertical = cs.verticalAlign || (header ? 'middle' : 'top');
  return {
    horizontal: /right/i.test(horizontal) ? 'right' : (/center|middle/i.test(horizontal) ? 'center' : 'left'),
    vertical: /bottom/i.test(vertical) ? 'bottom' : (/middle|center/i.test(vertical) ? 'middle' : 'top'),
    wrapText: cs.whiteSpace !== 'nowrap'
  };
}

function buildBorder(cs) {
  const top = parseBorderSide(cs, 'Top');
  const right = parseBorderSide(cs, 'Right');
  const bottom = parseBorderSide(cs, 'Bottom');
  const left = parseBorderSide(cs, 'Left');
  if (!top && !right && !bottom && !left) return defaultBorder();
  return { top, right, bottom, left };
}

function defaultBorder() {
  return {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } }
  };
}

function parseBorderSide(cs, side) {
  const style = cs[`border${side}Style`];
  const width = parseFloat(cs[`border${side}Width`]) || 0;
  const color = cs[`border${side}Color`];
  if (!style || style === 'none' || style === 'hidden' || width === 0) return undefined;
  const excelStyle = width >= 2 ? 'medium' : (BORDER_STYLE_MAP[style] || 'thin');
  return {
    style: excelStyle,
    color: { argb: cssColorToARGB(color, 'FF000000') }
  };
}

function buildFill(cs, header) {
  const bg = cs.backgroundColor;
  if (!bg || bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent') {
    return header ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } } : undefined;
  }
  return {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: cssColorToARGB(bg, 'FFFFFFFF') }
  };
}

function parseCssWidth(el) {
  const styleWidth = el.style?.width;
  if (!styleWidth) return null;
  const parsed = parseFloat(styleWidth);
  return Number.isNaN(parsed) ? null : parsed;
}
