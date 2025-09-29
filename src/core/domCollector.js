const DEFAULT_SELECTORS = [
  { type: 'letterhead', selector: '[data-export-block="letterhead"], .letterhead' },
  { type: 'info-grid', selector: '[data-export-block="info"], .info' },
  { type: 'remarks', selector: '[data-export-block="remarks"], .remarks' },
  { type: 'table', selector: '[data-export-block="table"], table' },
  { type: 'signature', selector: '[data-export-block="signature"], .signature' },
  { type: 'footer', selector: '[data-export-block="footer"], .footer' },
  { type: 'note', selector: '[data-export-block="note"], .note' }
];

export function collectBlocks(root, options = {}) {
  if (!root) return [];
  const selectors = options.selectors || DEFAULT_SELECTORS;
  const queue = [];
  selectors.forEach(cfg => {
    const elements = Array.from(root.querySelectorAll(cfg.selector));
    elements.forEach(el => {
      if (!isVisible(el)) return;
      queue.push({ type: cfg.type, element: el, weight: cfg.weight || 0 });
    });
  });

  const uniqueEntries = new Map();
  queue.forEach(entry => {
    const key = entry.element;
    if (!uniqueEntries.has(key) || uniqueEntries.get(key).weight < entry.weight) {
      uniqueEntries.set(key, entry);
    }
  });

  return Array.from(uniqueEntries.values()).sort((a, b) => {
    if (a.element === b.element) return 0;
    const pos = a.element.compareDocumentPosition(b.element);
    if (pos & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
    if (pos & Node.DOCUMENT_POSITION_PRECEDING) return 1;
    return 0;
  });
}

function isVisible(el) {
  const cs = getComputedStyle(el);
  if (cs.display === 'none' || cs.visibility === 'hidden') return false;
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}
