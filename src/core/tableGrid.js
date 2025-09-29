import { buildColumnWidthFromElement, extractCellStyle } from '../utils/style.js';

function parseSpan(cell, attr) {
  const span = parseInt(cell.getAttribute(attr), 10);
  return Number.isNaN(span) || span < 1 ? 1 : span;
}

function detectValue(cell) {
  const img = cell.querySelector('img');
  if (img) {
    return { type: 'image', element: cell, img };
  }
  const link = cell.querySelector('a[href]');
  const text = cell.textContent.trim();
  if (link && link.href) {
    return { type: 'link', text: link.textContent.trim() || link.href, href: link.href };
  }
  const numericMatch = text.replace(/[,\s]/g, '');
  if (/^-?\d+(\.\d+)?$/.test(numericMatch)) {
    return { type: 'number', text, numeric: Number(numericMatch) };
  }
  return { type: 'text', text };
}

export function buildColumnDefinitions(table) {
  const headerRow = table.querySelector('thead tr');
  const colgroup = table.querySelectorAll('colgroup col');
  if (colgroup.length > 0) {
    return Array.from(colgroup, col => {
      const width = col.style?.width || col.getAttribute('width');
      if (width) {
        const temp = document.createElement('div');
        temp.style.width = width;
        document.body.appendChild(temp);
        const rect = temp.getBoundingClientRect();
        temp.remove();
        return { width: rect.width ? Math.max(8, Math.round(((rect.width - 12) / 7) * 100) / 100) : 18 };
      }
      return { width: 18 };
    });
  }
  if (!headerRow) {
    return [];
  }
  return Array.from(headerRow.children).map(th => buildColumnWidthFromElement(th));
}

export function extractTable(table) {
  const rows = [];
  const merges = [];
  const occupancy = [];
  const sections = ['thead', 'tbody', 'tfoot'];
  sections.forEach(section => {
    const sectionEl = table.querySelector(section);
    if (!sectionEl) return;
    const trList = Array.from(sectionEl.querySelectorAll(':scope > tr'));
    trList.forEach(tr => {
      const rowIndex = rows.length;
      occupancy[rowIndex] = occupancy[rowIndex] || [];
      const excelRow = {
        section,
        cells: []
      };
      let colCursor = 0;
      const cells = Array.from(tr.children);
      cells.forEach(cell => {
        while (occupancy[rowIndex][colCursor]) {
          colCursor += 1;
        }
        const colspan = parseSpan(cell, 'colspan');
        const rowspan = parseSpan(cell, 'rowspan');
        const colStart = colCursor;
        const colEnd = colCursor + colspan - 1;
        const rowStart = rowIndex;
        const rowEnd = rowIndex + rowspan - 1;

        for (let r = rowStart; r <= rowEnd; r += 1) {
          occupancy[r] = occupancy[r] || [];
          for (let c = colStart; c <= colEnd; c += 1) {
            occupancy[r][c] = true;
          }
        }

        const value = detectValue(cell);
        const style = extractCellStyle(cell, { header: cell.tagName === 'TH' });

        excelRow.cells.push({
          column: colStart,
          colspan,
          rowspan,
          element: cell,
          style,
          value,
          isHeader: cell.tagName === 'TH'
        });

        if (colspan > 1 || rowspan > 1) {
          merges.push({
            startRow: rowIndex + 1,
            startCol: colStart + 1,
            endRow: rowIndex + rowspan,
            endCol: colStart + colspan
          });
        }

        colCursor = colEnd + 1;
      });
      rows.push(excelRow);
    });
  });

  const columnCount = occupancy.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);

  return {
    columnCount,
    rows,
    merges
  };
}
