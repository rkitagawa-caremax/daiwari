const waitForNextPaint = () => new Promise((resolve) => {
  requestAnimationFrame(() => requestAnimationFrame(resolve));
});

const waitForImage = async (image) => {
  if (!image) return;
  if (!image.complete) {
    await new Promise((resolve) => {
      const finish = () => resolve();
      image.addEventListener('load', finish, { once: true });
      image.addEventListener('error', finish, { once: true });
    });
  }
  if (typeof image.decode === 'function') {
    try {
      await image.decode();
    } catch {
      // load/error 待機済みのため、decode 非対応・失敗時もキャプチャを続行する。
    }
  }
};

export const waitForPdfExportSurface = async (sheetId, timeoutMs = 8000) => {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    await waitForNextPaint();
    const surface = Array.from(document.querySelectorAll('[data-pdf-export-sheet-id]'))
      .find((element) => element.dataset.pdfExportSheetId === String(sheetId));
    if (surface) {
      if (document.fonts?.ready) await document.fonts.ready;
      await Promise.all(Array.from(surface.querySelectorAll('img')).map(waitForImage));
      return surface;
    }
  }
  throw new Error('PDF出力用ページの描画がタイムアウトしました。');
};

export const createPdfRenderer = async () => {
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import('html2canvas-pro'),
    import('jspdf')
  ]);
  const documentPdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
    compress: true,
    putOnlyUsedFonts: true
  });
  let pageCount = 0;

  return {
    async appendSurface(surface) {
      const canvas = await html2canvas(surface, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        allowTaint: false,
        logging: false,
        imageTimeout: 20000,
        removeContainer: true
      });
      const imageData = canvas.toDataURL('image/jpeg', 0.94);
      if (pageCount > 0) documentPdf.addPage('a4', 'portrait');
      documentPdf.addImage(imageData, 'JPEG', 0, 0, 210, 297, undefined, 'FAST');
      pageCount += 1;
      canvas.width = 1;
      canvas.height = 1;
    },
    save(filename) {
      if (pageCount === 0) throw new Error('PDFへ追加されたページがありません。');
      documentPdf.save(filename);
    }
  };
};
