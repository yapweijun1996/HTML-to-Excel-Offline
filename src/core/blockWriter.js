import { extractTable } from './tableGrid.js';
import { measureElement, ensureRowHeight, pxToExcelColWidth } from '../utils/measurement.js';

export class WorksheetComposer {
  constructor(sheet, imageManager, options = {}) {
    this.sheet = sheet;
    this.images = imageManager;
    this.row = 1;
    this.options = { spacingPx: 12, ...options };
    const columnWidthsPx = options.columnWidthsPx || [70, 150, 360, 110, 110, 150];
    this.sheet.columns = columnWidthsPx.map(px => ({ width: pxToExcelColWidth(px) }));
    this.defaultColumnCount = this.sheet.columns.length;
  }

  get columnCount() {
    const defined = Array.isArray(this.sheet.columns) ? this.sheet.columns.length : 0;
    const used = typeof this.sheet.columnCount === 'number' ? this.sheet.columnCount : 0;
    return Math.max(defined, used, this.defaultColumnCount || 1);
  }

  addSpacer(px = this.options.spacingPx) {
    this.row += 1;
    ensureRowHeight(this.sheet, this.row, px);
  }

  writeLetterhead(element) {
    if (!element) return;
    const { sheet } = this;
    const logo = element.querySelector('img');
    const texts = Array.from(element.querySelectorAll('h1, h2, h3, p, .code, .sub'))
      .filter(el => el.textContent.trim());

    const lastColumn = Math.max(3, this.columnCount);
    sheet.mergeCells(this.row, 1, this.row + 1, 2);
    sheet.mergeCells(this.row, 3, this.row, lastColumn);
    const titleCell = sheet.getCell(this.row, 3);
    titleCell.value = texts.length ? texts[0].textContent.trim() : element.textContent.trim();
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { vertical: 'middle' };

    if (texts.length > 1) {
      sheet.mergeCells(this.row + 1, 3, this.row + 1, lastColumn);
      const subCell = sheet.getCell(this.row + 1, 3);
      subCell.value = texts.slice(1).map(el => el.textContent.trim()).join('\n');
      subCell.alignment = { wrapText: true };
      subCell.font = { size: 11 };
    }

    if (logo && logo.src) {
      const { rect } = measureElement(logo);
      this.images.queue({
        src: logo.src,
        sheet: sheet.name,
        row: this.row,
        col: 1,
        width: Math.max(60, rect.width || 120),
        height: Math.max(40, rect.height || 60),
        offsetCol: 0.05,
        offsetRow: 0.05
      });
      ensureRowHeight(sheet, this.row, rect.height + 16);
    }
    this.row += 2;
    this.addSpacer(6);
  }

  writeInfoGrid(element) {
    if (!element) return;
    const rows = collectInfoPairs(element);
    rows.forEach(pair => {
      this.row += 1;
      this.sheet.getCell(this.row, 1).value = pair.label;
      this.sheet.getCell(this.row, 1).font = { bold: true };
      this.sheet.mergeCells(this.row, 2, this.row, 3);
      this.sheet.getCell(this.row, 2).value = pair.value;
      this.sheet.getCell(this.row, 2).alignment = { wrapText: true };
      if (pair.label2) {
        this.sheet.getCell(this.row, 4).value = pair.label2;
        this.sheet.getCell(this.row, 4).font = { bold: true };
        this.sheet.mergeCells(this.row, 5, this.row, this.columnCount);
        this.sheet.getCell(this.row, 5).value = pair.value2;
        this.sheet.getCell(this.row, 5).alignment = { wrapText: true };
      }
      ensureRowHeight(this.sheet, this.row, pair.heightPx || 24);
    });
    this.addSpacer();
  }

  writeRemarks(element) {
    if (!element) return;
    this.row += 1;
    this.sheet.mergeCells(this.row, 1, this.row, this.columnCount);
    const titleCell = this.sheet.getCell(this.row, 1);
    titleCell.value = element.querySelector('h3')?.textContent.trim() || 'Remarks';
    titleCell.font = { bold: true };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCFCFC' } };

    const text = element.querySelector('p')?.textContent.trim() || element.textContent.trim();
    this.row += 1;
    this.sheet.mergeCells(this.row, 1, this.row, this.columnCount);
    const cell = this.sheet.getCell(this.row, 1);
    cell.value = text;
    cell.alignment = { wrapText: true };
    ensureRowHeight(this.sheet, this.row, measureElement(element).height || 36);
    this.addSpacer();
  }

  writeTable(element) {
    if (!element) return;

    const { rows, merges } = extractTable(element);
    const tableStartRow = this.row + 1;
    rows.forEach(dataRow => {
      this.row += 1;
      const rowNumber = this.row;
      const excelRow = this.sheet.getRow(rowNumber);
      dataRow.cells.forEach(cellData => {
        const columnIndex = cellData.column + 1;
        const cell = excelRow.getCell(columnIndex);
        applyCellValue(cell, cellData, this.images, this.sheet, rowNumber);
        applyCellStyle(cell, cellData);
        ensureRowHeight(this.sheet, rowNumber, measureElement(cellData.element).height || 24);
      });
    });

    merges.forEach(({ startRow, startCol, endRow, endCol }) => {
      const actualStart = tableStartRow + startRow - 1;
      const actualEnd = tableStartRow + endRow - 1;
      this.sheet.mergeCells(actualStart, startCol, actualEnd, endCol);
    });

    this.addSpacer();
  }

  writeNote(element) {
    if (!element) return;
    this.row += 1;
    this.sheet.mergeCells(this.row, 1, this.row, this.columnCount);
    const cell = this.sheet.getCell(this.row, 1);
    cell.value = element.textContent.trim();
    cell.font = { italic: true, color: { argb: 'FF555555' } };
    cell.alignment = { wrapText: true };
    ensureRowHeight(this.sheet, this.row, measureElement(element).height || 20);
  }

  writeFooter(element) {
    if (!element) return;
    this.row += 1;
    const textEl = element.querySelector('.footnote, [data-export-footnote]');
    if (textEl) {
      this.sheet.mergeCells(this.row, 1, this.row, this.columnCount - 1);
      const cell = this.sheet.getCell(this.row, 1);
      cell.value = textEl.textContent.trim();
      cell.alignment = { wrapText: true };
      ensureRowHeight(this.sheet, this.row, measureElement(textEl).height || 20);
    }
    const logo = element.querySelector('img');
    if (logo && logo.src) {
      const { rect } = measureElement(logo);
      this.images.queue({
        src: logo.src,
        sheet: this.sheet.name,
        row: this.row,
        col: this.columnCount,
        width: Math.max(40, rect.width || 80),
        height: Math.max(24, rect.height || 40),
        offsetCol: 0.1,
        offsetRow: 0.1
      });
      ensureRowHeight(this.sheet, this.row, rect.height + 12);
    }
    this.addSpacer();
  }
}

function collectInfoPairs(element) {
  const items = [];
  const keys = element.querySelectorAll('.k');
  if (keys.length) {
    for (let i = 0; i < keys.length; i += 2) {
      const label = keys[i];
      const value = label.nextElementSibling;
      const label2 = keys[i + 1];
      const value2 = label2 ? label2.nextElementSibling : null;
      items.push({
        label: label?.textContent.trim() || '',
        value: value?.textContent.trim() || '',
        label2: label2?.textContent.trim() || '',
        value2: value2?.textContent.trim() || '',
        heightPx: Math.max(measureElement(label).height, measureElement(value).height)
      });
    }
    return items;
  }
  const children = Array.from(element.children).filter(el => el.textContent.trim());
  for (let i = 0; i < children.length; i += 2) {
    items.push({
      label: children[i]?.textContent.trim() || '',
      value: children[i + 1]?.textContent.trim() || '',
      heightPx: Math.max(measureElement(children[i]).height, measureElement(children[i + 1]).height)
    });
  }
  return items;
}

function applyCellValue(cell, cellData, imageManager, sheet, rowNumber) {
  const { value, element, column } = cellData;
  if (value.type === 'number') {
    cell.value = value.numeric;
    cell.numFmt = '0.00';
    return;
  }
  if (value.type === 'link') {
    cell.value = { text: value.text, hyperlink: value.href };
    return;
  }
  if (value.type === 'image') {
    const box = element.querySelector('.imgbox') || element;
    const { rect } = measureElement(box);
    const width = Math.max(24, rect.width || 120);
    const height = Math.max(24, rect.height || 120);
    imageManager.queue({
      src: value.img.src,
      sheet: sheet.name,
      row: rowNumber,
      col: column + 1,
      width,
      height,
      offsetCol: 0.05,
      offsetRow: 0.05
    });
    cell.value = '';
    ensureRowHeight(sheet, rowNumber, height + 12);
    return;
  }
  cell.value = value.text || '';
}

function applyCellStyle(cell, cellData) {
  if (!cellData.style) return;
  const { alignment, font, border, fill } = cellData.style;
  if (alignment) cell.alignment = alignment;
  if (font) cell.font = font;
  if (border) cell.border = border;
  if (fill) cell.fill = fill;
}
