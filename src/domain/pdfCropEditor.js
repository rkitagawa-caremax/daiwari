// プレビュー上で切り抜き枠を手で調整するための純粋ロジック。
// 矩形はページ全体を 0〜1 とした正規化座標で、ページの外へは出さない。
// 手動で決めた枠はページ (batchPage.id) とコマ (row.id) の二段のマップで保持する。

export const MIN_PDF_CROP_SIZE = 0.015;
export const PDF_CROP_RESIZE_HANDLES = Object.freeze(['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se']);
export const EMPTY_PDF_CROP_MANUAL_RECTS = Object.freeze({});

const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
// px から求めた差分をそのまま足すと誤差が溜まるので、保存前に丸める
const round = (value) => Math.round(value * 1e6) / 1e6;

// 表示領域にページ全体が収まる大きさ (縦横比は保つ)
export const fitPdfPreviewSize = (page, stage) => {
  const pageWidth = Number(page?.width) || 0;
  const pageHeight = Number(page?.height) || 0;
  const stageWidth = Number(stage?.width) || 0;
  const stageHeight = Number(stage?.height) || 0;
  if (pageWidth <= 0 || pageHeight <= 0 || stageWidth <= 0 || stageHeight <= 0) return { width: 0, height: 0 };
  const scale = Math.min(stageWidth / pageWidth, stageHeight / pageHeight);
  return {
    width: Math.max(1, Math.floor(pageWidth * scale)),
    height: Math.max(1, Math.floor(pageHeight * scale))
  };
};

// 表示倍率 (1 = 表示領域にちょうど収まる大きさ)
export const PDF_PREVIEW_ZOOM_MIN = 1;
export const PDF_PREVIEW_ZOOM_MAX = 4;
export const PDF_PREVIEW_ZOOM_STEP = 0.25;

export const clampPdfPreviewZoom = (value) => {
  const zoom = Number(value);
  if (!Number.isFinite(zoom)) return PDF_PREVIEW_ZOOM_MIN;
  return Math.min(PDF_PREVIEW_ZOOM_MAX, Math.max(PDF_PREVIEW_ZOOM_MIN, Math.round(zoom * 100) / 100));
};

export const zoomPdfPreviewSize = (fit, zoom) => {
  if (!(fit?.width > 0) || !(fit?.height > 0)) return { width: 0, height: 0 };
  const scale = clampPdfPreviewZoom(zoom);
  return { width: Math.round(fit.width * scale), height: Math.round(fit.height * scale) };
};

// 枠ごと平行移動する (サイズは変えず、ページ内に収める)
export const movePdfCropRect = (rect, dx, dy) => ({
  x: round(clamp(rect.x + dx, 0, Math.max(0, 1 - rect.width))),
  y: round(clamp(rect.y + dy, 0, Math.max(0, 1 - rect.height))),
  width: rect.width,
  height: rect.height
});

// つまんだ辺 / 角だけを動かす (handle は 'nw' や 's' のような向き)
export const resizePdfCropRect = (rect, handle, dx, dy, minSize = MIN_PDF_CROP_SIZE) => {
  let left = rect.x;
  let top = rect.y;
  let right = rect.x + rect.width;
  let bottom = rect.y + rect.height;
  if (handle.includes('w')) left = clamp(left + dx, 0, Math.max(0, right - minSize));
  if (handle.includes('e')) right = clamp(right + dx, Math.min(1, left + minSize), 1);
  if (handle.includes('n')) top = clamp(top + dy, 0, Math.max(0, bottom - minSize));
  if (handle.includes('s')) bottom = clamp(bottom + dy, Math.min(1, top + minSize), 1);
  return { x: round(left), y: round(top), width: round(right - left), height: round(bottom - top) };
};

export const applyPdfCropDrag = (rect, handle, dx, dy) => (
  handle === 'move' ? movePdfCropRect(rect, dx, dy) : resizePdfCropRect(rect, handle, dx, dy)
);

export const getPdfCropManualRects = (overrides, pageId) => overrides?.[pageId] || EMPTY_PDF_CROP_MANUAL_RECTS;

export const setPdfCropManualRect = (overrides = {}, pageId, rowId, rect) => {
  if (!pageId || !rowId || !rect) return overrides;
  return { ...overrides, [pageId]: { ...getPdfCropManualRects(overrides, pageId), [rowId]: rect } };
};

export const clearPdfCropManualRect = (overrides = {}, pageId, rowId) => {
  const page = overrides?.[pageId];
  if (!page || !(rowId in page)) return overrides;
  const next = { ...page };
  delete next[rowId];
  return { ...overrides, [pageId]: next };
};

export const clearPdfCropManualPage = (overrides = {}, pageId) => {
  if (!overrides?.[pageId]) return overrides;
  const next = { ...overrides };
  delete next[pageId];
  return next;
};

export const countPdfCropManualRects = (overrides, pageId) => (
  pageId === undefined
    ? Object.values(overrides || {}).reduce((total, page) => total + Object.keys(page || {}).length, 0)
    : Object.keys(getPdfCropManualRects(overrides, pageId)).length
);
