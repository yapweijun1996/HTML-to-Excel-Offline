import { pxToExcelColWidth, pxToPt } from '../utils/measurement.js';
import { urlToDataURL, dataUrlToExtension } from '../utils/image.js';
import { cssColorToARGB } from '../utils/style.js';

export async function writeLayoutToWorkbook(workbook, analysis, options = {}) {
  const sheet = workbook.addWorksheet(options.sheetName || 'Exported Layout', {
    properties: { defaultRowHeight: pxToPt(analysis.rowHeights?.[0] || 20) }
  });
  applyColumns(sheet, analysis);
  applyRows(sheet, analysis);
  const occupancy = createOccupancyMatrix(analysis);
  const imageJobs = [];
  const sortedCells = [...analysis.cells].sort((a, b) => area(a) - area(b));
  sortedCells.forEach(cell => {
    if (isOccupied(occupancy, cell)) return;
    occupy(occupancy, cell);
    const excelCell = sheet.getCell(cell.rowStart + 1, cell.colStart + 1);
    if (cell.colEnd > cell.colStart || cell.rowEnd > cell.rowStart) {
      sheet.mergeCells(
        cell.rowStart + 1,
        cell.colStart + 1,
        cell.rowEnd + 1,
        cell.colEnd + 1
      );
    }
    applyCellContent(excelCell, cell, imageJobs, sheet);
    applyCellStyle(excelCell, cell);
  });

  await embedImages(workbook, imageJobs);
  return sheet;
}

function applyColumns(sheet, analysis) {
  const widths = analysis.columnWidths || [];
  if (!widths.length) return;
  sheet.columns = widths.map(px => ({ width: pxToExcelColWidth(px) }));
}

function applyRows(sheet, analysis) {
  const heights = analysis.rowHeights || [];
  heights.forEach((px, index) => {
    if (px <= 0) return;
    sheet.getRow(index + 1).height = pxToPt(px);
  });
}

function createOccupancyMatrix(analysis) {
  const cols = (analysis.columnGuides?.length || 1) - 1;
  const rows = (analysis.rowGuides?.length || 1) - 1;
  return Array.from({ length: rows }, () => Array(cols).fill(false));
}

function isOccupied(matrix, cell) {
  for (let r = cell.rowStart; r <= cell.rowEnd; r += 1) {
    for (let c = cell.colStart; c <= cell.colEnd; c += 1) {
      if (matrix[r]?.[c]) return true;
    }
  }
  return false;
}

function occupy(matrix, cell) {
  for (let r = cell.rowStart; r <= cell.rowEnd; r += 1) {
    for (let c = cell.colStart; c <= cell.colEnd; c += 1) {
      if (matrix[r]) matrix[r][c] = true;
    }
  }
}

function area(cell) {
  return (cell.colEnd - cell.colStart + 1) * (cell.rowEnd - cell.rowStart + 1);
}

function applyCellContent(excelCell, cell, imageJobs, sheet) {
  if (cell.type === 'image') {
    const img = cell.element.tagName === 'IMG' ? cell.element : cell.element.querySelector('img');
    if (img && img.src) {
      imageJobs.push({
        src: img.src,
        sheetName: sheet.name,
        range: {
          col: cell.colStart,
          row: cell.rowStart,
          width: cell.bounds.width,
          height: cell.bounds.height
        }
      });
    }
    excelCell.value = '';
    return;
  }

  if (cell.text) {
    excelCell.value = cell.text;
  } else {
    excelCell.value = '';
  }
}

function applyCellStyle(excelCell, cell) {
 if (cell.bounds?.isBold) {
    excelCell.font = { ...(excelCell.font || {}), bold: true };
  }
  excelCell.alignment = {
    horizontal: normalizeHorizontal(cell.bounds?.textAlign),
    vertical: 'top',
    wrapText: true
  };

  if (cell.bounds?.backgroundColor && cell.bounds.backgroundColor !== 'rgba(0, 0, 0, 0)') {
    excelCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cssColorToARGB(cell.bounds.backgroundColor) }
    };
  }
}

function normalizeHorizontal(textAlign) {
  if (!textAlign) return 'left';
  if (/center/i.test(textAlign)) return 'center';
  if (/right/i.test(textAlign)) return 'right';
  return 'left';
}

async function embedImages(workbook, jobs) {
  if (!jobs.length) return;
  for (const job of jobs) {
    const dataUrl = await urlToDataURL(job.src);
    if (!dataUrl) continue;
    const imgId = workbook.addImage({
      base64: dataUrl,
      extension: dataUrlToExtension(dataUrl)
    });
    const sheet = workbook.getWorksheet(job.sheetName);
    if (!sheet) continue;
    sheet.addImage(imgId, {
      tl: { col: job.range.col, row: job.range.row },
      ext: { width: job.range.width, height: job.range.height }
    });
  }
}
