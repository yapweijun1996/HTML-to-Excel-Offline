export const PX_PER_INCH = 96;
const CM_PER_INCH = 2.54;

export const pxToPt = px => Math.round((px * 72 / PX_PER_INCH) * 100) / 100;
export const pxToExcelColWidth = px => Math.max(6, Math.round(((px - 12) / 7) * 100) / 100);

export function parseCssLength(value, contextPx = 0) {
  if (!value) return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (/^-?\d+(?:\.\d+)?px$/.test(trimmed)) {
    return parseFloat(trimmed);
  }
  if (/^-?\d+(?:\.\d+)?cm$/.test(trimmed)) {
    return parseFloat(trimmed) * PX_PER_INCH / CM_PER_INCH;
  }
  if (/^-?\d+(?:\.\d+)?mm$/.test(trimmed)) {
    return parseFloat(trimmed) * PX_PER_INCH / (CM_PER_INCH * 10);
  }
  if (/^-?\d+(?:\.\d+)?in$/.test(trimmed)) {
    return parseFloat(trimmed) * PX_PER_INCH;
  }
  if (/^-?\d+(?:\.\d+)?pt$/.test(trimmed)) {
    return parseFloat(trimmed) * PX_PER_INCH / 72;
  }
  if (/^-?\d+(?:\.\d+)?%$/.test(trimmed)) {
    const percent = parseFloat(trimmed);
    return contextPx ? contextPx * percent / 100 : null;
  }
  return null;
}

export function measureElement(el) {
  if (!el) return { width: 0, height: 0, rect: { width: 0, height: 0 } };
  const rect = el.getBoundingClientRect();
  const computed = getComputedStyle(el);
  const padding = ['paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft'].reduce((acc, prop) => {
    acc[prop] = parseFloat(computed[prop]) || 0;
    return acc;
  }, {});
  return {
    width: rect.width,
    height: rect.height,
    rect,
    padding,
    computed
  };
}

export function ensureRowHeight(sheet, rowNumber, targetPx) {
  if (!targetPx) return;
  const current = sheet.getRow(rowNumber).height || 0;
  const required = pxToPt(targetPx);
  sheet.getRow(rowNumber).height = Math.max(current, required);
}
