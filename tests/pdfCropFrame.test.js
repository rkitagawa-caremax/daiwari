import test from 'node:test';
import assert from 'node:assert/strict';

import {
  chooseFrameEdge,
  computeDarkCoverageProfiles,
  countSnappedEdges,
  detectFrameLines,
  snapRectToFrame
} from '../src/domain/pdfCropFrame.js';

// 行 coverage を「区間 → 値」で組み立てるヘルパー
const buildCoverage = (length, segments, base = 0) => {
  const coverage = new Float32Array(length).fill(base);
  segments.forEach(([start, end, value]) => {
    for (let index = start; index <= end; index++) coverage[index] = value;
  });
  return coverage;
};

test('computeDarkCoverageProfiles counts dark pixels per row and column', () => {
  // 4x3 の画像: 1 行目は全て黒、2 行目は左端だけ黒、3 行目は白
  const width = 4;
  const height = 3;
  const data = new Uint8ClampedArray(width * height * 4).fill(255);
  const paintBlack = (x, y) => {
    const index = (y * width + x) * 4;
    data[index] = 0; data[index + 1] = 0; data[index + 2] = 0;
  };
  for (let x = 0; x < width; x++) paintBlack(x, 0);
  paintBlack(0, 1);

  const { rows, cols } = computeDarkCoverageProfiles(data, width, height);
  assert.deepEqual(Array.from(rows), [1, 0.25, 0]);
  assert.deepEqual(Array.from(cols).map((value) => Number(value.toFixed(4))), [0.6667, 0.3333, 0.3333, 0.3333]);
});

test('detectFrameLines keeps thin high-coverage runs and drops thick bands', () => {
  const coverage = buildCoverage(100, [
    [10, 12, 0.9],   // 枠線 (3px)
    [40, 60, 0.95],  // 商品画像や色帯 (21px) → 除外
    [80, 80, 0.7]    // 枠線 (1px)
  ]);
  assert.deepEqual(detectFrameLines(coverage, { minCoverage: 0.55, maxThickness: 6 }), [
    { start: 10, end: 12, center: 11 },
    { start: 80, end: 80, center: 80 }
  ]);
});

test('chooseFrameEdge picks the inner line of a neighbouring border pair', () => {
  // 上のコマの下枠 (40-42) と、余白 (43-52) を挟んだ自コマの上枠 (53-55)
  const coverage = buildCoverage(300, [
    [40, 42, 0.9],
    [53, 55, 0.9],
    [56, 200, 0.3]   // コマ内部 (文字などで薄く暗い)
  ]);
  const lines = detectFrameLines(coverage, { minCoverage: 0.55, maxThickness: 6 });
  // 期待位置 (45) は上のコマの下枠に近いが、外側が白い自コマの上枠 53 が選ばれる
  assert.equal(chooseFrameEdge(lines, coverage, { expected: 45, side: 'start', tolerance: 30, maxPairGap: 15, whiteMax: 0.03 }), 53);
  // 下辺: 自コマの下枠 (end) は外側 (下) が白い
  const bottomCoverage = buildCoverage(300, [
    [200, 202, 0.9],
    [214, 216, 0.9],
    [217, 299, 0.3]
  ]);
  const bottomLines = detectFrameLines(bottomCoverage, { minCoverage: 0.55, maxThickness: 6 });
  assert.equal(chooseFrameEdge(bottomLines, bottomCoverage, { expected: 210, side: 'end', tolerance: 30, maxPairGap: 15, whiteMax: 0.03 }), 202);
  assert.equal(chooseFrameEdge(bottomLines, bottomCoverage, { expected: 100, side: 'end', tolerance: 10, maxPairGap: 15, whiteMax: 0.03 }), null);
});

test('snapRectToFrame moves the grid rect onto the detected frame and reports snapped edges', () => {
  // 縦: 上のコマの下枠 40-42 / 余白 / 自コマ 53-55 〜 240-242 / 余白 / 次のコマ 252-254
  const rows = buildCoverage(300, [
    [40, 42, 0.9], [53, 55, 0.9], [56, 239, 0.3], [240, 242, 0.9], [252, 254, 0.9]
  ]);
  // 横: 左枠 20-22, 右枠 220-222
  const cols = buildCoverage(260, [[20, 22, 0.9], [23, 219, 0.3], [220, 222, 0.9]]);
  const expected = { x: 24, y: 45, width: 190, height: 187 }; // グリッド計算がやや上にずれている想定

  const snapped = snapRectToFrame({ profiles: { rows, cols }, expected });
  assert.deepEqual(snapped, {
    x: 19, y: 52, width: 204, height: 191,
    snappedEdges: { top: true, bottom: true, left: true, right: true }
  });
  assert.equal(countSnappedEdges(snapped.snappedEdges), 4);
});

test('snapRectToFrame falls back to the expected rect when no frame is found or the size is implausible', () => {
  const blankRows = buildCoverage(300, []);
  const blankCols = buildCoverage(260, []);
  const expected = { x: 24, y: 45, width: 190, height: 187 };
  const untouched = snapRectToFrame({ profiles: { rows: blankRows, cols: blankCols }, expected });
  assert.deepEqual(untouched, {
    ...expected,
    snappedEdges: { top: false, bottom: false, left: false, right: false }
  });

  // 上下とも許容範囲ぎりぎりの線が見つかるが、採用すると高さが 30% 以上縮む → 縦方向は期待値のまま
  const rows = buildCoverage(300, [[88, 90, 0.9], [190, 192, 0.9]]);
  const partial = snapRectToFrame({ profiles: { rows, cols: blankCols }, expected });
  assert.equal(partial.y, expected.y);
  assert.equal(partial.height, expected.height);
  assert.equal(partial.snappedEdges.top, false);
});
