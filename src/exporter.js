import { collectBlocks } from './core/domCollector.js';
import { WorksheetComposer } from './core/blockWriter.js';
import { ImageManager } from './core/imageManager.js';

const DEFAULT_OPTIONS = {
  mode: 'structure',
  sheetName: 'Export',
  defaultColumnCount: 6
};

function saveAsBlob(blob, filename) {
  const link = document.createElement('a');
  link.download = filename;
  link.href = URL.createObjectURL(blob);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function setStatus(statusEl, message) {
  if (statusEl) statusEl.textContent = message || '';
}

async function ensureImagesLoaded(container) {
  const imgs = Array.from(container.querySelectorAll('img'));
  if (!imgs.length) return;
  await Promise.all(imgs.map(img => new Promise(resolve => {
    if (img.complete && img.naturalWidth) return resolve();
    img.addEventListener('load', resolve, { once: true });
    img.addEventListener('error', resolve, { once: true });
    setTimeout(resolve, 2000);
  })));
}

async function buildWorkbook(root, options) {
  const ExcelJS = window.ExcelJS;
  if (!ExcelJS) {
    throw new Error('ExcelJS 未加载');
  }
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HTML to Excel Engine';
  wb.created = new Date();
  const sheet = wb.addWorksheet(options.sheetName || 'Export', {
    properties: { defaultRowHeight: 20 }
  });

  const imageManager = new ImageManager();
  const composer = new WorksheetComposer(sheet, imageManager, {
    defaultColumnCount: options.defaultColumnCount
  });

  const blocks = collectBlocks(root, options.blockSelectorOptions);
  if (!blocks.length) {
    const table = root.querySelector('table');
    if (table) {
      composer.writeTable(table);
    }
  } else {
    blocks.forEach(block => {
      switch (block.type) {
        case 'letterhead':
          composer.writeLetterhead(block.element);
          break;
        case 'info-grid':
          composer.writeInfoGrid(block.element);
          break;
        case 'remarks':
          composer.writeRemarks(block.element);
          break;
        case 'table':
          composer.writeTable(block.element);
          break;
        case 'footer':
          composer.writeFooter(block.element);
          break;
        case 'note':
          composer.writeNote(block.element);
          break;
        default:
          // fallback: 将文本写入单行
          composer.writeNote(block.element);
      }
    });
  }

  return { workbook: wb, imageManager };
}

async function exportHtml(options = {}) {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const exportBtn = document.getElementById(opts.buttonId || 'exportBtn');
  const statusEl = document.getElementById(opts.statusId || 'status');
  const scope = document.querySelector(opts.scopeSelector || 'body');
  if (!scope) throw new Error('找不到导出范围');

  if (exportBtn) exportBtn.disabled = true;
  setStatus(statusEl, '准备图片...');
  await ensureImagesLoaded(scope);

  setStatus(statusEl, '构建工作簿...');
  const { workbook, imageManager } = await buildWorkbook(scope.querySelector(opts.containerSelector || '.a4') || scope, opts);

  setStatus(statusEl, '嵌入图片...');
  const blocked = await imageManager.embedAll(workbook);

  setStatus(statusEl, '写入 .xlsx ...');
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAsBlob(blob, opts.filename || 'export.xlsx');

  if (blocked) {
    setStatus(statusEl, `完成，但 ${blocked} 张图片被 CORS 拒绝`);
  } else {
    setStatus(statusEl, '导出完成');
  }
  if (exportBtn) exportBtn.disabled = false;
}

function bindButton(options = {}) {
  const btn = document.getElementById(options.buttonId || 'exportBtn');
  if (!btn) return;
  btn.addEventListener('click', () => {
    exportHtml(options).catch(err => {
      console.error(err);
      const statusEl = document.getElementById(options.statusId || 'status');
      setStatus(statusEl, `导出失败: ${err.message}`);
      btn.disabled = false;
    });
  });
}

window.HtmlToExcelExporter = {
  export: exportHtml,
  bindButton
};

bindButton();
