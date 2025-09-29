const DEFAULT_OPTIONS = {
  blockDefinitions: null,
  includeGenericText: true
};

const DEFAULT_BLOCK_DEFS = [
  { type: 'letterhead', selector: '[data-export-block="letterhead"], .letterhead, header[role="banner"]' },
  { type: 'info-grid', selector: '[data-export-block="info"], .info, dl[data-export-block], dl.info, dl.kv' },
  { type: 'remarks', selector: '[data-export-block="remarks"], .remarks, blockquote[data-type="remarks"]' },
  { type: 'signature', selector: '[data-export-block="signature"], .signature' },
  { type: 'footer', selector: '[data-export-block="footer"], footer, .footer, [role="contentinfo"]' },
  { type: 'note', selector: '[data-export-block="note"], .note, aside[data-role="note"]' },
  { type: 'table', selector: '[data-export-block="table"], table' }
];

export function collectBlocks(root, options = {}) {
  if (!root) return [];
  const mergedOptions = { ...DEFAULT_OPTIONS, ...options };
  const defs = mergedOptions.blockDefinitions || DEFAULT_BLOCK_DEFS;
  const visited = mergedOptions.visited || new Set();
  const blocks = [];

  const pushBlock = (type, element, meta = {}) => {
    if (!element || visited.has(element)) return;
    if (!isElementVisible(element)) return;
    visited.add(element);
    blocks.push({ type, element, meta });
  };

  // 1. definitions based pass
  defs.forEach(def => {
    const matches = resolveMatches(root, def);
    matches.forEach(el => pushBlock(def.type, el, def.meta || {}));
  });

  // 2. fallback discovery via traversal in DOM order
  const traverse = node => {
    if (!node || visited.has(node)) return;
    if (!isElementVisible(node)) return;
    if (node.closest('[data-export-skip="true"]')) return;

    const classification = classifyNode(node);
    if (classification) {
      pushBlock(classification.type, node, classification.meta);
      return; // stop at this node to avoid duplicate nested matches
    }

    Array.from(node.children).forEach(child => traverse(child));
  };

  Array.from(root.children).forEach(child => traverse(child));

  // 3. fallback for text nodes if requested
  if (mergedOptions.includeGenericText) {
    const textBlocks = detectStandaloneTextBlocks(root, visited);
    textBlocks.forEach(block => pushBlock('text', block.element, block.meta));
  }

  // sort blocks by document order
  blocks.sort((a, b) => compareDocumentOrder(a.element, b.element));
  return blocks;
}

function resolveMatches(root, def) {
  if (!def) return [];
  if (typeof def.find === 'function') {
    const result = def.find(root) || [];
    return Array.isArray(result) ? result : [result];
  }
  if (def.selector) {
    return Array.from(root.querySelectorAll(def.selector));
  }
  return [];
}

function classifyNode(node) {
  if (!node || node.nodeType !== Node.ELEMENT_NODE) return null;
  const explicit = node.getAttribute('data-export-block');
  if (explicit) {
    return { type: explicit.trim(), meta: {} };
  }
  if (node.tagName === 'TABLE') {
    return { type: 'table', meta: {} };
  }
  if (/^H[1-6]$/.test(node.tagName)) {
    const level = Number(node.tagName.slice(1));
    return { type: 'text', meta: { headingLevel: level } };
  }
  if (node.tagName === 'DL') {
    return { type: 'info-grid', meta: {} };
  }
  if (node.matches('.letterhead, header[role="banner"]')) {
    return { type: 'letterhead', meta: {} };
  }
  if (node.matches('.remarks, blockquote[data-type="remarks"]')) {
    return { type: 'remarks', meta: {} };
  }
  if (node.matches('footer, .footer, [role="contentinfo"]')) {
    return { type: 'footer', meta: {} };
  }
  if (node.matches('aside, .note, [data-role="note"]')) {
    return { type: 'note', meta: {} };
  }
  if (node.matches('.signature, [data-signature]')) {
    return { type: 'signature', meta: {} };
  }
  if (node.matches('section[data-export="text"]')) {
    return { type: 'text', meta: {} };
  }
  return null;
}

function detectStandaloneTextBlocks(root, visited) {
  const blocks = [];
  const candidates = Array.from(root.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, blockquote'));
  candidates.forEach(node => {
    if (visited.has(node)) return;
    if (hasVisitedAncestor(node, visited)) return;
    if (!isElementVisible(node)) return;
    if (node.closest('[data-export-skip="true"]')) return;
    if (node.closest('table')) return;
    if (!hasMeaningfulText(node)) return;
    if (blocks.some(entry => entry.element.contains(node))) return;
    blocks.push({ element: node, meta: deriveTextMeta(node) });
  });
  return blocks;
}

function deriveTextMeta(node) {
  if (/^H[1-6]$/.test(node.tagName)) {
    return { headingLevel: Number(node.tagName.slice(1)) };
  }
  if (node.tagName === 'LI') {
    return { addSpacing: false };
  }
  if (node.matches('p[align="center"], .text-center, [data-align="center"]')) {
    return { align: 'center' };
  }
  if (node.matches('p[align="right"], .text-right, [data-align="right"]')) {
    return { align: 'right' };
  }
  return {};
}

function compareDocumentOrder(a, b) {
  if (a === b) return 0;
  const pos = a.compareDocumentPosition(b);
  if (pos & Node.DOCUMENT_POSITION_PRECEDING) return 1;
  if (pos & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
  return 0;
}

function isElementVisible(el) {
  if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
  const style = getComputedStyle(el);
  if (style.display === 'none' || style.visibility === 'hidden') return false;
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

function hasMeaningfulText(el) {
  const text = el.innerText || '';
  return text.trim().length > 0;
}

function isTextualTag(el) {
  return ['P', 'SPAN', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE'].includes(el.tagName);
}

function hasVisitedAncestor(node, visited) {
  let current = node.parentElement;
  while (current) {
    if (visited.has(current)) return true;
    current = current.parentElement;
  }
  return false;
}
