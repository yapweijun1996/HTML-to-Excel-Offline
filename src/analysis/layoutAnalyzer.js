/**
 * Layout Analyzer
 * ----------------
 * Given a DOM container, measure all visible elements and projects them onto a
 * normalized grid so Excel can mirror the HTML layout without template-specific
 * logic.
 */

import { measureElement } from '../utils/measurement.js';

const DEFAULT_OPTIONS = {
  tolerancePx: 2,
  minLineDistancePx: 4,
  minCellWidthPx: 12,
  minCellHeightPx: 12,
  includeImages: true,
  includeTextNodes: true
};

export function analyzeLayout(container, options = {}) {
  if (!container) throw new Error('Layout analyzer: container is required');
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const containerRect = container.getBoundingClientRect();
  const elements = collectElements(container, containerRect, opts);
  const guides = buildGuides(elements, containerRect, opts);
  const cells = mapElementsToGrid(elements, guides, opts);
  return {
    ...guides,
    elements,
    cells
  };
}

function collectElements(container, containerRect, options) {
  const items = [];
  const walker = document.createTreeWalker(
    container,
    NodeFilter.SHOW_ELEMENT,
    {
      acceptNode(node) {
        if (!isVisible(node)) return NodeFilter.FILTER_REJECT;
        if (node.closest('[data-export-skip="true"]')) return NodeFilter.FILTER_REJECT;
        if (node === container) return NodeFilter.FILTER_ACCEPT;
        return NodeFilter.FILTER_ACCEPT;
      }
    }
  );

  let current = walker.currentNode;
  if (current === container) current = walker.nextNode();
  while (current) {
    if (current !== container) {
      const info = describeElement(current, containerRect, options);
      if (info) items.push(info);
    }
    current = walker.nextNode();
  }
  return items;
}

function describeElement(element, containerRect, options) {
  const rect = element.getBoundingClientRect();
  if (!rect || rect.width < 1 || rect.height < 1) return null;
  const relative = {
    left: rect.left - containerRect.left,
    top: rect.top - containerRect.top,
    right: rect.right - containerRect.left,
    bottom: rect.bottom - containerRect.top,
    width: rect.width,
    height: rect.height
  };
  const styles = getComputedStyle(element);
  const info = {
    element,
    rect,
    relative,
    left: relative.left,
    top: relative.top,
    right: relative.right,
    bottom: relative.bottom,
    width: relative.width,
    height: relative.height,
    tag: element.tagName,
    display: styles.display,
    position: styles.position,
    zIndex: parseInt(styles.zIndex, 10) || 0,
    fontWeight: styles.fontWeight,
    fontStyle: styles.fontStyle,
    textAlign: styles.textAlign,
    verticalAlign: styles.verticalAlign,
    whiteSpace: styles.whiteSpace,
    backgroundColor: styles.backgroundColor,
    color: styles.color,
    borderTop: styles.borderTop,
    borderRight: styles.borderRight,
    borderBottom: styles.borderBottom,
    borderLeft: styles.borderLeft,
    isBold: isBold(styles)
  };

  if (element.tagName === 'TABLE') {
    if (element === element.closest('table')) return null; // skip table container
    info.type = 'table';
  } else if (element.tagName === 'TR') {
    return null;
  } else if (element.tagName === 'IMG' || element.querySelector('img')) {
    info.type = 'image';
  } else if (element.tagName === 'TD' || element.tagName === 'TH') {
    info.type = 'table-cell';
  } else {
    const text = element.textContent.trim();
    info.text = text;
    info.type = text ? 'text' : 'block';
  }
  return info;
}

function buildGuides(elements, containerRect, options) {
  const tolerance = options.tolerancePx;
  const columns = [];
  const rows = [];
  pushGuide(columns, 0, tolerance);
  pushGuide(rows, 0, tolerance);
  const containerWidth = containerRect?.width || 0;
  const containerHeight = containerRect?.height || 0;
  if (containerWidth) pushGuide(columns, containerWidth, tolerance);
  if (containerHeight) pushGuide(rows, containerHeight, tolerance);
  elements.forEach(item => {
    pushGuide(columns, item.left, tolerance);
    pushGuide(columns, item.right, tolerance);
    pushGuide(rows, item.top, tolerance);
    pushGuide(rows, item.bottom, tolerance);
  });
  columns.sort((a, b) => a - b);
  rows.sort((a, b) => a - b);
  const columnGuides = normalizeGuides(columns, tolerance);
  const rowGuides = normalizeGuides(rows, tolerance);
  return {
    columnGuides,
    rowGuides,
    columnWidths: diffValues(columnGuides),
    rowHeights: diffValues(rowGuides)
  };
}

function mapElementsToGrid(elements, guides, options) {
  const cells = [];
  const colGuides = guides.columnGuides;
  const rowGuides = guides.rowGuides;
  elements.forEach(item => {
    const colStart = findGuideIndex(colGuides, item.left, options.tolerancePx);
    const colEnd = findGuideIndex(colGuides, item.right, options.tolerancePx) - 1;
    const rowStart = findGuideIndex(rowGuides, item.top, options.tolerancePx);
    const rowEnd = findGuideIndex(rowGuides, item.bottom, options.tolerancePx) - 1;
    if (colStart == null || colEnd == null || rowStart == null || rowEnd == null) return;
    if (colEnd < colStart || rowEnd < rowStart) return;
    cells.push({
      element: item.element,
      type: item.type,
      text: item.text,
      bounds: item,
      colStart,
      colEnd,
      rowStart,
      rowEnd
    });
  });
  return cells;
}

function isVisible(el) {
  const styles = getComputedStyle(el);
  if (styles.display === 'none' || styles.visibility === 'hidden' || styles.opacity === '0') return false;
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

function pushGuide(guides, value, tolerance) {
  for (let i = 0; i < guides.length; i += 1) {
    if (Math.abs(guides[i] - value) <= tolerance) {
      guides[i] = (guides[i] + value) / 2;
      return;
    }
  }
  guides.push(value);
}

function findGuideIndex(guides, value, tolerance) {
  for (let i = 0; i < guides.length; i += 1) {
    if (Math.abs(guides[i] - value) <= tolerance) return i;
    if (value < guides[i]) return i;
  }
  return guides.length ? guides.length - 1 : null;
}

function normalizeGuides(values, tolerance) {
  if (!values.length) return [];
  values.sort((a, b) => a - b);
  const result = [values[0]];
  for (let i = 1; i < values.length; i += 1) {
    const value = values[i];
    const last = result[result.length - 1];
    if (Math.abs(last - value) > tolerance) {
      result.push(value);
    } else {
      result[result.length - 1] = (last + value) / 2;
    }
  }
  return result;
}

function diffValues(values) {
  const diffs = [];
  for (let i = 1; i < values.length; i += 1) {
    diffs.push(values[i] - values[i - 1]);
  }
  return diffs;
}

function isBold(styles) {
  const weight = parseInt(styles.fontWeight, 10);
  if (!Number.isNaN(weight)) return weight >= 600;
  return styles.fontWeight && styles.fontWeight.toLowerCase().includes('bold');
}
