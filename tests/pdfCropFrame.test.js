import test from 'node:test';
import assert from 'node:assert/strict';

import {
  chooseGutterEdge,
  chooseLineEdge,
  computeCoverageProfiles,
  countSnappedEdges,
  detectFrameLines,
  findWhiteRuns,
  snapRectToFrame
} from '../src/domain/pdfCropFrame.js';

// coverage を「区間 → 値」で組み立てるヘルパー (指定外は base)
const buildCoverage = (length, segments, base = 0) => {
  const coverage = new Float32Array(length).fill(base);
  segments.forEach(([start, end, value]) => {
    for (let index = start; index <= end; index++) coverage[index] = value;
  });
  return coverage;
};

test('computeCoverageProfiles measures rows over the core columns only (and vice versa)', () => {
  // 6x4 の画像。左端 1 列は隣のコマの文字 (黒)、行 1 は全面黒、それ以外は白
  const width = 6;
  const height = 4;
  const data = new Uint8ClampedArray(width * height * 4).fill(255);
  const paintBlack = (x, y) => {
    const index = (y * width + x) * 4;
    data[index] = 0; data[index + 1] = 0; data[index + 2] = 0;
  };
  for (let y = 0; y < height; y++) paintBlack(0, y);
  for (let x = 0; x < width; x++) paintBlack(x, 1);

  const { rowsInk, colsInk, rowsDark } = computeCoverageProfiles(data, width, height, { core: { x0: 1, x1: 6, y0: 0, y1: 4 } });
  // 行: 隣コマ (x=0) は集計外なので、黒い行 1 以外は 0
  assert.deepEqual(Array.from(rowsInk), [0, 1, 0, 0]);
  assert.deepEqual(Array.from(rowsDark), [0, 1, 0, 0]);
  // 列: x=0 は全行黒、その他は行 1 だけ黒
  assert.deepEqual(Array.from(colsInk), [1, 0.25, 0.25, 0.25, 0.25, 0.25]);
});

test('findWhiteRuns returns runs of near-white entries at least minLength long', () => {
  const coverage = buildCoverage(40, [[0, 4, 0], [5, 9, 0.5], [10, 21, 0.01], [22, 30, 0.4], [31, 32, 0], [33, 39, 0.6]], 0.5);
  assert.deepEqual(findWhiteRuns(coverage, { whiteMax: 0.02, minLength: 3 }), [
    { start: 0, end: 4, length: 5 },
    { start: 10, end: 21, length: 12 }
  ]);
});

test('detectFrameLines keeps thin high-coverage runs and drops thick bands', () => {
  const coverage = buildCoverage(100, [
    [10, 12, 0.9],
    [40, 60, 0.95],
    [80, 80, 0.7]
  ]);
  assert.deepEqual(detectFrameLines(coverage, { minCoverage: 0.55, maxThickness: 6 }), [
    { start: 10, end: 12, center: 11 },
    { start: 80, end: 80, center: 80 }
  ]);
});

test('chooseGutterEdge picks the gutter nearest the expected edge and pads into it', () => {
  const runs = [
    { start: 40, end: 52, length: 13 },   // 上のコマとの余白
    { start: 120, end: 160, length: 41 }, // コマ内部の広い白帯 (タイトルと画像の間)
    { start: 240, end: 251, length: 12 }  // 下のコマとの余白
  ];
  // 上辺: 期待 56 → 余白 40-52 の内側 53 から 4px 余白側へ
  assert.equal(chooseGutterEdge(runs, { expected: 56, side: 'start', tolerance: 25, pad: 4 }), 49);
  // 下辺: 期待 236 → 余白 240-251 の内側 239 から 4px 余白側へ
  assert.equal(chooseGutterEdge(runs, { expected: 236, side: 'end', tolerance: 25, pad: 4 }), 243);
  // 許容範囲外なら null
  assert.equal(chooseGutterEdge(runs, { expected: 200, side: 'end', tolerance: 10, pad: 4 }), null);
  // パディングは余白の半分まで
  assert.equal(chooseGutterEdge([{ start: 40, end: 43, length: 4 }], { expected: 44, side: 'start', tolerance: 10, pad: 10 }), 42);
});

test('chooseLineEdge returns the inside of the nearest rule so the rule is not cropped in', () => {
  const lines = [{ start: 40, end: 42, center: 41 }, { start: 250, end: 252, center: 251 }];
  assert.equal(chooseLineEdge(lines, { expected: 45, side: 'start', tolerance: 10 }), 43);
  assert.equal(chooseLineEdge(lines, { expected: 245, side: 'end', tolerance: 10 }), 249);
  // パディングを付けるとさらに内側へ
  assert.equal(chooseLineEdge(lines, { expected: 45, side: 'start', tolerance: 10, pad: 2 }), 45);
  assert.equal(chooseLineEdge(lines, { expected: 245, side: 'end', tolerance: 10, pad: 2 }), 247);
  // 遠い線は使わない
  assert.equal(chooseLineEdge(lines, { expected: 150, side: 'end', tolerance: 10 }), null);
});

test('snapRectToFrame prefers the ruled lines: thick rows, dotted columns', () => {
  // 縦: 行の区切りは太い実線 (48-53 と 244-249)。すぐ上には余白もあるが、罫線を優先する
  const rowsDark = buildCoverage(300, [[48, 53, 0.9], [244, 249, 0.9]], 0.1);
  const rowsInk = buildCoverage(300, [[36, 47, 0]], 0.3);
  // 横: 列の区切りは点線 (18-19 と 232-233)。途切れるので濃さは 0.35 しかない
  const colsDark = buildCoverage(260, [[18, 19, 0.35], [232, 233, 0.35]], 0.05);
  const colsInk = buildCoverage(260, [], 0.3);
  const expected = { x: 24, y: 45, width: 200, height: 190 };

  const snapped = snapRectToFrame({ profiles: { rowsInk, rowsDark, colsInk, colsDark }, expected });
  assert.deepEqual(snapped, {
    // 上 54+1、下 243-1（+1 で end 243）、左 20+1、右 231-1（+1 で end 231）
    x: 21,
    y: 55,
    width: 210,
    height: 188,
    snappedEdges: { top: true, bottom: true, left: true, right: true }
  });
  assert.equal(countSnappedEdges(snapped.snappedEdges), 4);
});

test('snapRectToFrame falls back to gutters on edges without a rule', () => {
  // 縦は罫線あり、横は罫線が無く余白だけ
  const rowsDark = buildCoverage(300, [[48, 53, 0.9], [244, 249, 0.9]], 0.1);
  const rowsInk = buildCoverage(300, [], 0.3);
  const colsDark = buildCoverage(260, [], 0.05);
  const colsInk = buildCoverage(260, [[15, 26, 0], [223, 234, 0]], 0.3);
  const expected = { x: 24, y: 45, width: 200, height: 190 };

  const snapped = snapRectToFrame({ profiles: { rowsInk, rowsDark, colsInk, colsDark }, expected });
  assert.deepEqual(snapped.snappedEdges, { top: true, bottom: true, left: true, right: true });
  assert.equal(snapped.y, 55, '上下は罫線の内側');
  // 左右は余白の内側 + パディング（200*0.02=4）
  assert.equal(snapped.x, 23);
  assert.equal(snapped.x + snapped.width, 227);
});

test('snapRectToFrame falls back to the expected rect when nothing plausible is found', () => {
  const expected = { x: 24, y: 45, width: 190, height: 187 };
  const blank = { rowsInk: buildCoverage(300, [], 0.3), rowsDark: buildCoverage(300, []), colsInk: buildCoverage(260, [], 0.3), colsDark: buildCoverage(260, []) };
  assert.deepEqual(snapRectToFrame({ profiles: blank, expected }), {
    ...expected,
    snappedEdges: { top: false, bottom: false, left: false, right: false }
  });

  // 上下の余白が見つかるが、採用すると高さが 30% 以上縮む → 縦方向は期待値のまま
  const rowsInk = buildCoverage(300, [[65, 75, 0], [195, 205, 0]], 0.3);
  const partial = snapRectToFrame({ profiles: { ...blank, rowsInk }, expected, options: { searchToleranceRatio: 0.25 } });
  assert.equal(partial.y, expected.y);
  assert.equal(partial.height, expected.height);
  assert.equal(partial.snappedEdges.top, false);
});
