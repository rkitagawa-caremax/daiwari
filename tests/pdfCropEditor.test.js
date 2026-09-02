import test from 'node:test';
import assert from 'node:assert/strict';

import {
  applyPdfCropDrag,
  clampPdfPreviewZoom,
  clearPdfCropManualPage,
  clearPdfCropManualRect,
  countPdfCropManualRects,
  fitPdfPreviewSize,
  getPdfCropManualRects,
  MIN_PDF_CROP_SIZE,
  movePdfCropRect,
  PDF_PREVIEW_ZOOM_MAX,
  PDF_PREVIEW_ZOOM_MIN,
  resizePdfCropRect,
  setPdfCropManualRect,
  setPdfCropManualRects,
  zoomPdfPreviewSize
} from '../src/domain/pdfCropEditor.js';

const RECT = { x: 0.2, y: 0.3, width: 0.25, height: 0.2 };

test('fitPdfPreviewSize scales the page to fill the stage without distorting it', () => {
  // 横長の表示領域 → 高さが上限
  assert.deepEqual(fitPdfPreviewSize({ width: 800, height: 1000 }, { width: 900, height: 500 }), { width: 400, height: 500 });
  // 縦長の表示領域 → 幅が上限
  assert.deepEqual(fitPdfPreviewSize({ width: 800, height: 1000 }, { width: 400, height: 900 }), { width: 400, height: 500 });
  // 測定前 / 未描画は 0
  assert.deepEqual(fitPdfPreviewSize({ width: 800, height: 1000 }, { width: 0, height: 0 }), { width: 0, height: 0 });
  assert.deepEqual(fitPdfPreviewSize(null, { width: 400, height: 900 }), { width: 0, height: 0 });
});

test('preview zoom stays within range and scales the fitted size', () => {
  assert.equal(clampPdfPreviewZoom(2.5), 2.5);
  assert.equal(clampPdfPreviewZoom(0.2), PDF_PREVIEW_ZOOM_MIN, '等倍より小さくはしない');
  assert.equal(clampPdfPreviewZoom(99), PDF_PREVIEW_ZOOM_MAX);
  assert.equal(clampPdfPreviewZoom('x'), PDF_PREVIEW_ZOOM_MIN);
  assert.deepEqual(zoomPdfPreviewSize({ width: 400, height: 500 }, 1), { width: 400, height: 500 });
  assert.deepEqual(zoomPdfPreviewSize({ width: 400, height: 500 }, 2.5), { width: 1000, height: 1250 });
  assert.deepEqual(zoomPdfPreviewSize({ width: 0, height: 0 }, 2), { width: 0, height: 0 });
});

test('movePdfCropRect keeps the size and stops at the page edge', () => {
  assert.deepEqual(movePdfCropRect(RECT, 0.05, -0.1), { x: 0.25, y: 0.2, width: 0.25, height: 0.2 });
  assert.deepEqual(movePdfCropRect(RECT, -0.5, -0.5), { x: 0, y: 0, width: 0.25, height: 0.2 });
  assert.deepEqual(movePdfCropRect(RECT, 0.9, 0.9), { x: 0.75, y: 0.8, width: 0.25, height: 0.2 });
});

test('resizePdfCropRect moves only the grabbed edges', () => {
  // 下辺だけ下げる
  assert.deepEqual(resizePdfCropRect(RECT, 's', 0.3, 0.04), { x: 0.2, y: 0.3, width: 0.25, height: 0.24 });
  // 左上の角
  assert.deepEqual(resizePdfCropRect(RECT, 'nw', -0.05, -0.1), { x: 0.15, y: 0.2, width: 0.3, height: 0.3 });
  // 右下の角
  assert.deepEqual(resizePdfCropRect(RECT, 'se', 0.05, 0.05), { x: 0.2, y: 0.3, width: 0.3, height: 0.25 });
});

test('resizePdfCropRect stays inside the page and keeps a minimum size', () => {
  const wide = resizePdfCropRect(RECT, 'e', 0.9);
  assert.equal(wide.x + wide.width, 1);
  const tiny = resizePdfCropRect(RECT, 'w', 0.9, 0);
  assert.equal(tiny.width, MIN_PDF_CROP_SIZE);
  assert.equal(tiny.x + tiny.width, RECT.x + RECT.width, '掴んでいない右辺は動かない');
  const flat = resizePdfCropRect(RECT, 'n', 0, 0.9);
  assert.equal(flat.height, MIN_PDF_CROP_SIZE);
});

test('applyPdfCropDrag routes move and resize handles', () => {
  assert.deepEqual(applyPdfCropDrag(RECT, 'move', 0.05, 0), movePdfCropRect(RECT, 0.05, 0));
  assert.deepEqual(applyPdfCropDrag(RECT, 'se', 0.05, 0.05), resizePdfCropRect(RECT, 'se', 0.05, 0.05));
});

test('manual rect overrides are stored per page and per frame', () => {
  const first = setPdfCropManualRect({}, 'page-1', 'row-a', RECT);
  const second = setPdfCropManualRect(first, 'page-1', 'row-b', RECT);
  const third = setPdfCropManualRect(second, 'page-2', 'row-a', RECT);
  assert.deepEqual(Object.keys(getPdfCropManualRects(third, 'page-1')), ['row-a', 'row-b']);
  assert.equal(countPdfCropManualRects(third), 3);
  assert.equal(countPdfCropManualRects(third, 'page-1'), 2);
  assert.equal(countPdfCropManualRects(third, 'page-9'), 0);

  const withoutRow = clearPdfCropManualRect(third, 'page-1', 'row-a');
  assert.deepEqual(Object.keys(getPdfCropManualRects(withoutRow, 'page-1')), ['row-b']);
  assert.equal(clearPdfCropManualRect(withoutRow, 'page-1', 'row-a'), withoutRow, '無い枠のクリアは同じ参照を返す');

  const withoutPage = clearPdfCropManualPage(third, 'page-1');
  assert.equal(countPdfCropManualRects(withoutPage), 1);
  assert.equal(clearPdfCropManualPage(withoutPage, 'page-1'), withoutPage);
  // 元の状態は書き換えない
  assert.equal(countPdfCropManualRects(third), 3);
});

test('setPdfCropManualRects writes several frames at once without touching others', () => {
  const base = setPdfCropManualRect({}, 'page-1', 'row-a', RECT);
  const moved = { ...RECT, x: 0.5 };
  const next = setPdfCropManualRects(base, 'page-1', { 'row-b': moved, 'row-c': moved });
  assert.deepEqual(Object.keys(getPdfCropManualRects(next, 'page-1')).sort(), ['row-a', 'row-b', 'row-c']);
  assert.deepEqual(getPdfCropManualRects(next, 'page-1')['row-a'], RECT);
  assert.deepEqual(getPdfCropManualRects(next, 'page-1')['row-b'], moved);
  // 空の更新やページ未指定は同じ参照を返す
  assert.equal(setPdfCropManualRects(base, 'page-1', {}), base);
  assert.equal(setPdfCropManualRects(base, undefined, { 'row-b': moved }), base);
});
