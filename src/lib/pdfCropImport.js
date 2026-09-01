import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

import { DEFAULT_FRAME_SNAP_OPTIONS, computeCoverageProfiles, countSnappedEdges, snapRectToFrame } from '../domain/pdfCropFrame.js';

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

// 枠線検出のために期待枠の周囲をどれだけ広く読むか (辺の長さに対する比率)
const FRAME_SEARCH_EXPANSION = 0.3;
const NO_SNAP = Object.freeze({ top: false, bottom: false, left: false, right: false });

export const openPdfFile = async (file) => {
  if (!file) throw new Error('PDFファイルが選択されていません。');
  const data = new Uint8Array(await file.arrayBuffer());
  const assetBaseUrl = new URL(import.meta.env.BASE_URL || '/', window.location.origin);
  return pdfjs.getDocument({
    data,
    cMapUrl: new URL('pdfjs/cmaps/', assetBaseUrl).toString(),
    cMapPacked: true,
    standardFontDataUrl: new URL('pdfjs/standard_fonts/', assetBaseUrl).toString(),
    useSystemFonts: true
  }).promise;
};

export const renderPdfPage = async (pdfDocument, pageNumber, { scale = 1.25, canvas } = {}) => {
  if (!pdfDocument) throw new Error('PDFが読み込まれていません。');
  const page = await pdfDocument.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const targetCanvas = canvas || document.createElement('canvas');
  const context = targetCanvas.getContext('2d', { alpha: false });
  targetCanvas.width = Math.ceil(viewport.width);
  targetCanvas.height = Math.ceil(viewport.height);

  await page.render({ canvas: targetCanvas, canvasContext: context, viewport }).promise;

  const textContent = await page.getTextContent();
  const textItems = textContent.items
    .filter((item) => item?.str)
    .map((item) => {
      // x, y は文字の左端 / ベースライン。width, height は viewport 尺に直してから正規化する
      const [x, y] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
      return {
        text: item.str,
        x: x / viewport.width,
        y: y / viewport.height,
        width: ((item.width || 0) * viewport.scale) / viewport.width,
        height: ((item.height || 0) * viewport.scale) / viewport.height
      };
    });

  return {
    canvas: targetCanvas,
    pageNumber,
    width: targetCanvas.width,
    height: targetCanvas.height,
    textItems
  };
};

const canvasToBlob = (canvas, type, quality) => new Promise((resolve, reject) => {
  canvas.toBlob((blob) => {
    if (blob) resolve(blob);
    else reject(new Error('切り抜き画像を生成できませんでした。'));
  }, type, quality);
});

export const cropPdfPageToFile = async ({
  canvas,
  normalizedRect,
  filename,
  maxSize = 900,
  quality = 0.92
}) => {
  if (!canvas || !normalizedRect) throw new Error('切り抜き範囲がありません。');

  const sourceX = Math.max(0, Math.round(normalizedRect.x * canvas.width));
  const sourceY = Math.max(0, Math.round(normalizedRect.y * canvas.height));
  const sourceWidth = Math.min(canvas.width - sourceX, Math.max(1, Math.round(normalizedRect.width * canvas.width)));
  const sourceHeight = Math.min(canvas.height - sourceY, Math.max(1, Math.round(normalizedRect.height * canvas.height)));
  const outputScale = Math.min(1, maxSize / Math.max(sourceWidth, sourceHeight));

  const outputCanvas = document.createElement('canvas');
  outputCanvas.width = Math.max(1, Math.round(sourceWidth * outputScale));
  outputCanvas.height = Math.max(1, Math.round(sourceHeight * outputScale));
  const outputContext = outputCanvas.getContext('2d', { alpha: false });
  outputContext.fillStyle = '#ffffff';
  outputContext.fillRect(0, 0, outputCanvas.width, outputCanvas.height);
  outputContext.drawImage(
    canvas,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    0,
    0,
    outputCanvas.width,
    outputCanvas.height
  );

  const blob = await canvasToBlob(outputCanvas, 'image/jpeg', quality);
  return new File([blob], filename, { type: 'image/jpeg', lastModified: Date.now() });
};

// 描画済みキャンバス上で、グリッドから求めた正規化矩形をコマの境界 (余白/枠線) にスナップさせる。
// 境界が見つからない場合は元の矩形をそのまま返す (snappedCount = 0)。
export const refineCropRectToFrame = (canvas, normalizedRect, options = {}) => {
  const fallback = { rect: normalizedRect, snappedEdges: NO_SNAP, snappedCount: 0 };
  if (!canvas || !normalizedRect || !canvas.width || !canvas.height) return fallback;

  const px = {
    x: normalizedRect.x * canvas.width,
    y: normalizedRect.y * canvas.height,
    width: normalizedRect.width * canvas.width,
    height: normalizedRect.height * canvas.height
  };
  const windowX = Math.max(0, Math.floor(px.x - px.width * FRAME_SEARCH_EXPANSION));
  const windowY = Math.max(0, Math.floor(px.y - px.height * FRAME_SEARCH_EXPANSION));
  const windowRight = Math.min(canvas.width, Math.ceil(px.x + px.width * (1 + FRAME_SEARCH_EXPANSION)));
  const windowBottom = Math.min(canvas.height, Math.ceil(px.y + px.height * (1 + FRAME_SEARCH_EXPANSION)));
  const windowWidth = windowRight - windowX;
  const windowHeight = windowBottom - windowY;
  if (windowWidth < 8 || windowHeight < 8) return fallback;

  let imageData;
  try {
    imageData = canvas.getContext('2d', { willReadFrequently: true }).getImageData(windowX, windowY, windowWidth, windowHeight);
  } catch (error) {
    console.warn('Frame detection skipped (getImageData failed):', error);
    return fallback;
  }

  const expected = { x: px.x - windowX, y: px.y - windowY, width: px.width, height: px.height };
  const inset = options.coreInsetRatio ?? DEFAULT_FRAME_SNAP_OPTIONS.coreInsetRatio;
  const profiles = computeCoverageProfiles(imageData.data, windowWidth, windowHeight, {
    // 隣のコマの内容を拾わないよう、期待枠の中央部分だけで行/列プロファイルを取る
    core: {
      x0: expected.x + expected.width * inset,
      x1: expected.x + expected.width * (1 - inset),
      y0: expected.y + expected.height * inset,
      y1: expected.y + expected.height * (1 - inset)
    },
    inkThreshold: options.inkThreshold,
    darkThreshold: options.darkThreshold
  });
  const snapped = snapRectToFrame({ profiles, expected, options });
  return {
    rect: {
      x: (windowX + snapped.x) / canvas.width,
      y: (windowY + snapped.y) / canvas.height,
      width: snapped.width / canvas.width,
      height: snapped.height / canvas.height
    },
    snappedEdges: snapped.snappedEdges,
    snappedCount: countSnappedEdges(snapped.snappedEdges)
  };
};
