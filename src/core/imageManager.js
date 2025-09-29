import { urlToDataURL, dataUrlToExtension } from '../utils/image.js';

export class ImageManager {
  constructor() {
    this.jobs = [];
    this.cache = new Map();
  }

  queue(options) {
    if (!options || !options.src) return;
    const key = `${options.src}|${options.sheet}|${options.row}|${options.col}`;
    this.jobs.push({ ...options, key });
  }

  async embedAll(workbook) {
    let blocked = 0;
    for (const job of this.jobs) {
      const dataUrl = await this.ensureDataURL(job.src);
      if (!dataUrl) {
        blocked += 1;
        continue;
      }
      const imgId = workbook.addImage({
        base64: dataUrl,
        extension: dataUrlToExtension(dataUrl)
      });
      const sheet = workbook.getWorksheet(job.sheet);
      if (!sheet) continue;
      sheet.addImage(imgId, {
        tl: {
          col: (job.col - 1) + (job.offsetCol || 0.1),
          row: (job.row - 1) + (job.offsetRow || 0.1)
        },
        ext: {
          width: job.width,
          height: job.height
        },
        editAs: 'oneCell'
      });
    }
    return blocked;
  }

  async ensureDataURL(src) {
    if (this.cache.has(src)) {
      return this.cache.get(src);
    }
    const dataUrl = await urlToDataURL(src);
    this.cache.set(src, dataUrl);
    return dataUrl;
  }
}
