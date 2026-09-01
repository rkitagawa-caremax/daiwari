// 描画済みページのピクセルから、コマの枠線 (上下左右) を検出して切り抜き枠をスナップさせる純粋ロジック。
// 入力は RGBA バイト列と期待枠 (グリッドから計算した枠) で、Canvas には依存しない (lib/pdfCropImport.js が橋渡しする)。
//
// 考え方:
//   - 行ごと / 列ごとに「暗いピクセルの割合」(coverage) を出す
//   - 割合が高く、かつ細い行の連なり = 枠線 (太い帯や商品画像は除外)
//   - 期待した辺の近くにある枠線のうち、外側が白い余白 (隣のコマとの間隔) になっているものを優先して選ぶ
//     → 隣のコマの枠線を誤って拾わない
//   - 見つからない / サイズが大きく外れる場合は期待枠をそのまま使う

export const DEFAULT_FRAME_SNAP_OPTIONS = Object.freeze({
  darkThreshold: 110,        // 輝度がこの値未満なら「暗いピクセル」
  minLineCoverage: 0.55,     // 枠線とみなす最小の暗ピクセル割合 (探索窓の幅に対する比率)
  maxLineThicknessRatio: 0.015, // 枠線の最大太さ (期待枠の辺の長さに対する比率)
  whiteMaxCoverage: 0.03,    // 余白とみなす最大の暗ピクセル割合
  maxPairGapRatio: 0.08,     // 隣接コマの枠線ペアとみなす最大の間隔 (期待枠の辺の長さに対する比率)
  searchToleranceRatio: 0.25, // 期待した辺からこの範囲 (辺の長さ比) で枠線を探す
  sizeToleranceRatio: 0.3,   // スナップ後の幅/高さが期待値から ±この比率を超えたら採用しない
  outerPadding: 1            // 枠線を含めるための外側パディング (px)
});

// RGBA バイト列から行ごと・列ごとの暗ピクセル割合を求める。
export const computeDarkCoverageProfiles = (data, width, height, darkThreshold = DEFAULT_FRAME_SNAP_OPTIONS.darkThreshold) => {
  const rowCounts = new Uint32Array(height);
  const colCounts = new Uint32Array(width);
  for (let y = 0; y < height; y++) {
    const rowOffset = y * width * 4;
    for (let x = 0; x < width; x++) {
      const index = rowOffset + x * 4;
      const luminance = 0.299 * data[index] + 0.587 * data[index + 1] + 0.114 * data[index + 2];
      if (luminance < darkThreshold) {
        rowCounts[y]++;
        colCounts[x]++;
      }
    }
  }
  const rows = new Float32Array(height);
  const cols = new Float32Array(width);
  for (let y = 0; y < height; y++) rows[y] = width > 0 ? rowCounts[y] / width : 0;
  for (let x = 0; x < width; x++) cols[x] = height > 0 ? colCounts[x] / height : 0;
  return { rows, cols };
};

// coverage 配列から「細い枠線」の候補 (連続区間) を抽出する。
export const detectFrameLines = (coverage, { minCoverage, maxThickness }) => {
  const lines = [];
  let start = -1;
  const flush = (end) => {
    if (start < 0) return;
    const thickness = end - start + 1;
    if (thickness <= maxThickness) lines.push({ start, end, center: (start + end) / 2 });
    start = -1;
  };
  for (let index = 0; index < coverage.length; index++) {
    if (coverage[index] >= minCoverage) {
      if (start < 0) start = index;
    } else {
      flush(index - 1);
    }
  }
  flush(coverage.length - 1);
  return lines;
};

const isWhiteRange = (coverage, from, to, whiteMax) => {
  for (let index = from; index <= to; index++) {
    if (coverage[index] > whiteMax) return false;
  }
  return true;
};

// 「短い白い間隔を挟んで並ぶ 2 本の枠線」= 隣接するコマ同士の枠線ペア。
// 上/左辺を探すときはペアの後ろ側 (自コマ側)、下/右辺を探すときはペアの前側を優先する。
export const markFrameLinePairs = (lines, coverage, { maxPairGap, whiteMax }) => lines.map((line, index) => {
  const previous = lines[index - 1];
  const next = lines[index + 1];
  const pairedWithPrevious = !!previous
    && line.start - previous.end - 1 <= maxPairGap
    && isWhiteRange(coverage, previous.end + 1, line.start - 1, whiteMax);
  const pairedWithNext = !!next
    && next.start - line.end - 1 <= maxPairGap
    && isWhiteRange(coverage, line.end + 1, next.start - 1, whiteMax);
  return { ...line, pairedWithPrevious, pairedWithNext };
});

// 期待位置に最も近い枠線の辺を選ぶ。side='start' は上/左辺、'end' は下/右辺。
export const chooseFrameEdge = (lines, coverage, { expected, side, tolerance, maxPairGap, whiteMax }) => {
  const edgeOf = (line) => (side === 'start' ? line.start : line.end);
  const marked = markFrameLinePairs(lines, coverage, { maxPairGap, whiteMax });
  const candidates = marked.filter((line) => Math.abs(edgeOf(line) - expected) <= tolerance);
  if (candidates.length === 0) return null;
  // 隣のコマ側の枠線 (上辺探索ならペアの前側、下辺探索ならペアの後ろ側) は後回しにする
  const preferred = candidates.filter((line) => (side === 'start' ? !line.pairedWithNext : !line.pairedWithPrevious));
  const pool = preferred.length > 0 ? preferred : candidates;
  pool.sort((left, right) => Math.abs(edgeOf(left) - expected) - Math.abs(edgeOf(right) - expected));
  return edgeOf(pool[0]);
};

const snapAxis = (coverage, expectedStart, expectedSize, options) => {
  const maxThickness = Math.max(2, Math.round(expectedSize * options.maxLineThicknessRatio));
  const tolerance = expectedSize * options.searchToleranceRatio;
  const maxPairGap = Math.max(2, Math.round(expectedSize * options.maxPairGapRatio));
  const lines = detectFrameLines(coverage, { minCoverage: options.minLineCoverage, maxThickness });
  const shared = { tolerance, maxPairGap, whiteMax: options.whiteMaxCoverage };
  const start = chooseFrameEdge(lines, coverage, { ...shared, expected: expectedStart, side: 'start' });
  const end = chooseFrameEdge(lines, coverage, { ...shared, expected: expectedStart + expectedSize, side: 'end' });

  let resolvedStart = start === null ? expectedStart : start - options.outerPadding;
  let resolvedEnd = end === null ? expectedStart + expectedSize : end + options.outerPadding;
  const size = resolvedEnd - resolvedStart;
  const withinTolerance = size > 0 && Math.abs(size - expectedSize) <= expectedSize * options.sizeToleranceRatio;
  if (!withinTolerance) {
    resolvedStart = expectedStart;
    resolvedEnd = expectedStart + expectedSize;
  }
  return {
    start: resolvedStart,
    size: resolvedEnd - resolvedStart,
    snappedStart: withinTolerance && start !== null,
    snappedEnd: withinTolerance && end !== null
  };
};

// 探索窓内のピクセル分布 (profiles) と期待枠 (窓内 px 座標) から、枠線にスナップした矩形を返す。
export const snapRectToFrame = ({ profiles, expected, options = {} }) => {
  const resolved = { ...DEFAULT_FRAME_SNAP_OPTIONS, ...options };
  const vertical = snapAxis(profiles.rows, expected.y, expected.height, resolved);
  const horizontal = snapAxis(profiles.cols, expected.x, expected.width, resolved);
  return {
    x: horizontal.start,
    y: vertical.start,
    width: horizontal.size,
    height: vertical.size,
    snappedEdges: {
      top: vertical.snappedStart,
      bottom: vertical.snappedEnd,
      left: horizontal.snappedStart,
      right: horizontal.snappedEnd
    }
  };
};

export const countSnappedEdges = (snappedEdges = {}) => Object.values(snappedEdges).filter(Boolean).length;
