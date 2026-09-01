import {
  DEFAULT_PDF_GRID_BOUNDS,
  getPdfCropRectFromGrid,
  getPdfGridGeometry,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode
} from './pdfCropImport.js';

// --- ページごとのグリッド校正 ---
// PDF の文字レイヤーにある各コマのコードラベル (例: 261-E1773) は、そのコマの左下付近に置かれる。
// CSV から分かる各コマの列/行位置とラベル座標を突き合わせ、そのページの実際のグリッド
// (原点・セル幅・セル高) を中央値ベースで推定する。ラベル位置が特殊な大コマがあっても外れ値として吸収される。

export const PDF_CROP_LABEL_PAD_X_RATIO = 0.06;  // コマ左端 → ラベル左端 (セル幅比)
export const PDF_CROP_LABEL_PAD_Y_RATIO = 0.13;  // ラベル基準線 → コマ下端 (セル高比)
const CELL_SIZE_TOLERANCE_RATIO = 0.15;
const ORIGIN_SHIFT_TOLERANCE_RATIO = 0.5;

const median = (values) => {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};

// 各コマについて、期待枠 (±半セル) の中でコードと一致する文字のうち左下に最も近いものをアンカーにする
const collectPdfCropAnchors = (rows, textItems, grid) => {
  const anchors = [];
  const marginX = grid.cellWidth * 0.5;
  const marginY = grid.cellHeight * 0.5;
  rows.forEach((row) => {
    if (!row?.code || !isPdfCropRowInsideGrid(row)) return;
    const rect = getPdfCropRectFromGrid(row, grid);
    const candidates = textItems.filter((item) => (
      normalizePdfCropCode(item.text) === row.code
      && item.x >= rect.x - marginX && item.x <= rect.x + rect.width + marginX
      && item.y >= rect.y - marginY && item.y <= rect.y + rect.height + marginY
    ));
    if (candidates.length === 0) return;
    const anchorX = rect.x;
    const anchorY = rect.y + rect.height;
    candidates.sort((left, right) => (
      Math.hypot(left.x - anchorX, left.y - anchorY) - Math.hypot(right.x - anchorX, right.y - anchorY)
    ));
    anchors.push({
      rowId: row.id,
      column: row.xPos - 1,
      bottomLine: row.yPos + row.rowSpan - 1,
      x: candidates[0].x,
      y: candidates[0].y
    });
  });
  return anchors;
};

// 異なる列 (行) にあるラベル同士の距離からセルサイズを推定する (ペアの中央値)。
const estimateCellSize = (anchors, indexKey, positionKey, fallback) => {
  const slopes = [];
  for (let i = 0; i < anchors.length; i++) {
    for (let j = i + 1; j < anchors.length; j++) {
      const indexDelta = anchors[j][indexKey] - anchors[i][indexKey];
      if (indexDelta === 0) continue;
      slopes.push((anchors[j][positionKey] - anchors[i][positionKey]) / indexDelta);
    }
  }
  if (slopes.length === 0) return fallback;
  const estimate = median(slopes);
  return Math.abs(estimate - fallback) <= fallback * CELL_SIZE_TOLERANCE_RATIO ? estimate : fallback;
};

export const calibratePdfCropGrid = ({ rows = [], textItems = [], bounds = DEFAULT_PDF_GRID_BOUNDS } = {}) => {
  const base = getPdfGridGeometry(bounds);
  const anchors = collectPdfCropAnchors(rows, Array.isArray(textItems) ? textItems : [], base);
  if (anchors.length === 0) return { ...base, anchorCount: 0, calibrated: false };

  const cellWidth = estimateCellSize(anchors, 'column', 'x', base.cellWidth);
  const cellHeight = estimateCellSize(anchors, 'bottomLine', 'y', base.cellHeight);
  let left = median(anchors.map((anchor) => anchor.x - anchor.column * cellWidth)) - cellWidth * PDF_CROP_LABEL_PAD_X_RATIO;
  // コマ下端の y = top + bottomLine * cellHeight (bottomLine = yPos + rowSpan - 1 はグリッド線の番号)
  let top = median(anchors.map((anchor) => anchor.y - anchor.bottomLine * cellHeight)) + cellHeight * PDF_CROP_LABEL_PAD_Y_RATIO;
  if (Math.abs(left - base.left) > cellWidth * ORIGIN_SHIFT_TOLERANCE_RATIO) left = base.left;
  if (Math.abs(top - base.top) > cellHeight * ORIGIN_SHIFT_TOLERANCE_RATIO) top = base.top;

  return { left, top, cellWidth, cellHeight, anchorCount: anchors.length, calibrated: true };
};
