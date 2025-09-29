const dataUrlRegex = /^data:image\/(png|jpe?g|gif|webp);base64,/i;

export async function urlToDataURL(url) {
  if (!url) return null;
  if (url.startsWith('data:')) {
    return await ensureExcelFriendly(url);
  }
  try {
    const resp = await fetch(url, { mode: 'cors', credentials: 'omit' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const blob = await resp.blob();
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
    return await ensureExcelFriendly(dataUrl);
  } catch (err) {
    console.warn('Image fetch blocked:', url, err.message);
    return null;
  }
}

async function ensureExcelFriendly(dataUrl) {
  if (!dataUrl) return null;
  if (dataUrl.startsWith('data:image/png') || dataUrl.startsWith('data:image/jpeg') || dataUrl.startsWith('data:image/jpg')) {
    return dataUrl;
  }
  if (dataUrl.startsWith('data:image/svg')) {
    return null; // TODO: 可扩展 SVG->PNG 转换
  }
  if (dataUrlRegex.test(dataUrl)) {
    return await rasterizeToPng(dataUrl);
  }
  return dataUrl;
}

async function rasterizeToPng(dataUrl) {
  return await new Promise(resolve => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth || img.width;
      canvas.height = img.naturalHeight || img.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(null);
    img.src = dataUrl;
  });
}

export function dataUrlToExtension(dataUrl) {
  if (!dataUrl) return 'png';
  if (dataUrl.startsWith('data:image/jpeg') || dataUrl.startsWith('data:image/jpg')) return 'jpeg';
  if (dataUrl.startsWith('data:image/gif')) return 'gif';
  if (dataUrl.startsWith('data:image/png')) return 'png';
  if (dataUrl.startsWith('data:image/webp')) return 'png';
  return 'png';
}
