import test from 'node:test';
import assert from 'node:assert/strict';

import {
  calibratePdfCropGrid,
  PDF_CROP_LABEL_PAD_X_RATIO,
  PDF_CROP_LABEL_PAD_Y_RATIO
} from '../src/domain/pdfCropGridCalibration.js';
import { getPdfCropRectFromGrid, getPdfGridGeometry } from '../src/domain/pdfCropImport.js';

// 実際の誌面グリッド (既定の余白設定とは少しずれている)
const TRUE_GRID = { left: 0.12, top: 0.08, cellWidth: 0.19, cellHeight: 0.21 };

const makeGridRows = () => {
  const rows = [];
  for (let y = 1; y <= 2; y++) {
    for (let x = 1; x <= 4; x++) {
      rows.push({ id: `r${y}${x}`, code: `E10${y}${x}`, xPos: x, yPos: y, rowSpan: 1, colSpan: 1, layoutStatus: 'ready' });
    }
  }
  return rows;
};

const makeLabelItems = (rows, grid = TRUE_GRID) => rows.map((row) => ({
  text: `261-${row.code}`,
  x: grid.left + (row.xPos - 1) * grid.cellWidth + grid.cellWidth * PDF_CROP_LABEL_PAD_X_RATIO,
  y: grid.top + (row.yPos + row.rowSpan - 1) * grid.cellHeight - grid.cellHeight * PDF_CROP_LABEL_PAD_Y_RATIO
}));

const near = (actual, expected, epsilon = 1e-6) => assert.ok(Math.abs(actual - expected) < epsilon, `${actual} is not close to ${expected}`);

test('calibratePdfCropGrid recovers the page grid from code label positions', () => {
  const rows = makeGridRows();
  const grid = calibratePdfCropGrid({ rows, textItems: [...makeLabelItems(rows), { text: '別の文字', x: 0.5, y: 0.5 }] });
  assert.equal(grid.calibrated, true);
  assert.equal(grid.anchorCount, 8);
  near(grid.left, TRUE_GRID.left);
  near(grid.top, TRUE_GRID.top);
  near(grid.cellWidth, TRUE_GRID.cellWidth);
  near(grid.cellHeight, TRUE_GRID.cellHeight);
  const rect = getPdfCropRectFromGrid(rows[5], grid); // yPos 2, xPos 2
  near(rect.x, TRUE_GRID.left + TRUE_GRID.cellWidth);
  near(rect.y, TRUE_GRID.top + TRUE_GRID.cellHeight);
});

test('calibratePdfCropGrid ignores an oddly placed label and falls back without anchors', () => {
  const rows = makeGridRows();
  const items = makeLabelItems(rows);
  items[3] = { ...items[3], x: items[3].x + 0.08, y: items[3].y - 0.12 }; // 大コマなどでラベルが右上にあるケース
  const grid = calibratePdfCropGrid({ rows, textItems: items });
  near(grid.left, TRUE_GRID.left, 1e-3);
  near(grid.top, TRUE_GRID.top, 1e-3);
  near(grid.cellWidth, TRUE_GRID.cellWidth, 1e-3);

  const fallback = calibratePdfCropGrid({ rows, textItems: [] });
  assert.equal(fallback.calibrated, false);
  assert.equal(fallback.anchorCount, 0);
  assert.deepEqual(
    { left: fallback.left, top: fallback.top, cellWidth: fallback.cellWidth, cellHeight: fallback.cellHeight },
    getPdfGridGeometry()
  );
});

test('calibratePdfCropGrid keeps the default cell width when only one column is available', () => {
  const rows = makeGridRows().filter((row) => row.xPos === 1);
  const grid = calibratePdfCropGrid({ rows, textItems: makeLabelItems(rows) });
  const base = getPdfGridGeometry();
  near(grid.cellWidth, base.cellWidth);
  near(grid.cellHeight, TRUE_GRID.cellHeight);
  // 左端はラベル位置から逆算 (パディングは既定セル幅基準になる分だけ僅かにずれる)
  near(grid.left, TRUE_GRID.left + TRUE_GRID.cellWidth * PDF_CROP_LABEL_PAD_X_RATIO - base.cellWidth * PDF_CROP_LABEL_PAD_X_RATIO);
});
