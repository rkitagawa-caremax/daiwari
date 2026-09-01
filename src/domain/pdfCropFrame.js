// 描画済みページのピクセルから、コマの境界 (余白 = ガター) を検出して切り抜き枠をスナップさせる純粋ロジック。
// 入力は RGBA バイト列と期待枠で、Canvas には依存しない (lib/pdfCropImport.js が橋渡しする)。
//
// 考え方:
//   - 期待枠の「中央部分」(core) だけを使って、行ごと / 列ごとの「インク (白でない) ピクセルの割合」を出す
//     → 隣のコマの内容が混ざりにくい
//   - 期待した辺の近くで、コマ幅いっぱいに白が続く帯 (余白) を探し、その内側の縁 + 少しの余白を境界にする
//     枠線がある誌面でも、枠線は余白の内側にあるインクとしてそのまま含まれる
//   - 余白が見つからない (コマ同士が線一本で接している) 場合は、細い濃い線 (枠線) の位置を使う
//   - どちらも見つからない / サイズが大きく外れる場合は期待枠をそのまま使う

export const DEFAULT_FRAME_SNAP_OPTIONS = Object.freeze({
  inkThreshold: 235,          // 輝度がこの値未満なら「インクあり」(白でない)
  darkThreshold: 150,         // 輝度がこの値未満なら「濃い」(枠線判定用)
  whiteMaxCoverage: 0.02,     // 余白とみなす最大のインク割合
  minGutterRatio: 0.006,      // 余白とみなす最小の連続長 (期待枠の辺の長さに対する比率)
  paddingRatio: 0.02,         // 境界を余白側へ広げる量 (辺の長さ比、余白の半分を上限)
  searchToleranceRatio: 0.12, // 期待した辺からこの範囲 (辺の長さ比) で余白/枠線を探す
  minLineCoverage: 0.55,      // 枠線とみなす最小の濃いピクセル割合
  maxLineThicknessRatio: 0.015, // 枠線の最大太さ (辺の長さ比)
  sizeToleranceRatio: 0.3,    // スナップ後の幅/高さが期待値から ±この比率を超えたら採用しない
  coreInsetRatio: 0.15        // プロファイル計算に使う中央部分 (両端をこの比率だけ除く)
});

const clampIndex = (value, max) => Math.min(max, Math.max(0, Math.round(value)));

// RGBA バイト列から行ごと・列ごとの「インク割合」と「濃いピクセル割合」を求める。
// 行プロファイルは core.x0〜x1 の列だけ、列プロファイルは core.y0〜y1 の行だけを集計する。
export const computeCoverageProfiles = (data, width, height, {
  core = { x0: 0, x1: width, y0: 0, y1: height },
  inkThreshold = DEFAULT_FRAME_SNAP_OPTIONS.inkThreshold,
  darkThreshold = DEFAULT_FRAME_SNAP_OPTIONS.darkThreshold
} = {}) => {
  const x0 = clampIndex(core.x0, width);
  const x1 = Math.max(x0 + 1, clampIndex(core.x1, width));
  const y0 = clampIndex(core.y0, height);
  const y1 = Math.max(y0 + 1, clampIndex(core.y1, height));
  const rowInk = new Uint32Array(height);
  const rowDark = new Uint32Array(height);
  const colInk = new Uint32Array(width);
  const colDark = new Uint32Array(width);

  for (let y = 0; y < height; y++) {
    const inCoreRows = y >= y0 && y < y1;
    const rowOffset = y * width * 4;
    for (let x = 0; x < width; x++) {
      const inCoreCols = x >= x0 && x < x1;
      if (!inCoreRows && !inCoreCols) continue;
      const index = rowOffset + x * 4;
      const luminance = 0.299 * data[index] + 0.587 * data[index + 1] + 0.114 * data[index + 2];
      const isInk = luminance < inkThreshold;
      const isDark = luminance < darkThreshold;
      if (inCoreCols) {
        if (isInk) rowInk[y]++;
        if (isDark) rowDark[y]++;
      }
      if (inCoreRows) {
        if (isInk) colInk[x]++;
        if (isDark) colDark[x]++;
      }
    }
  }

  const coreWidth = x1 - x0;
  const coreHeight = y1 - y0;
  const rowsInk = new Float32Array(height);
  const rowsDark = new Float32Array(height);
  const colsInk = new Float32Array(width);
  const colsDark = new Float32Array(width);
  for (let y = 0; y < height; y++) {
    rowsInk[y] = rowInk[y] / coreWidth;
    rowsDark[y] = rowDark[y] / coreWidth;
  }
  for (let x = 0; x < width; x++) {
    colsInk[x] = colInk[x] / coreHeight;
    colsDark[x] = colDark[x] / coreHeight;
  }
  return { rowsInk, rowsDark, colsInk, colsDark };
};

// coverage 配列から「白が連続する区間」(余白候補) を抽出する。
export const findWhiteRuns = (coverage, { whiteMax, minLength }) => {
  const runs = [];
  let start = -1;
  const flush = (end) => {
    if (start < 0) return;
    const length = end - start + 1;
    if (length >= minLength) runs.push({ start, end, length });
    start = -1;
  };
  for (let index = 0; index < coverage.length; index++) {
    if (coverage[index] <= whiteMax) {
      if (start < 0) start = index;
    } else {
      flush(index - 1);
    }
  }
  flush(coverage.length - 1);
  return runs;
};

// coverage 配列から「細い濃い線」(枠線候補) を抽出する。
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

// 期待位置に最も近い余白を選び、その内側の縁 (+余白側へのパディング) を境界として返す。
// side='start' は上/左辺 (余白の下/右端が内側)、'end' は下/右辺 (余白の上/左端が内側)。
export const chooseGutterEdge = (runs, { expected, side, tolerance, pad = 0 }) => {
  const candidates = runs
    .map((run) => {
      const innerEdge = side === 'start' ? run.end + 1 : run.start - 1;
      const padding = Math.min(pad, Math.floor(run.length / 2));
      return {
        innerEdge,
        boundary: side === 'start' ? innerEdge - padding : innerEdge + padding,
        distance: Math.abs(innerEdge - expected)
      };
    })
    .filter((candidate) => candidate.distance <= tolerance);
  if (candidates.length === 0) return null;
  candidates.sort((left, right) => left.distance - right.distance);
  return candidates[0].boundary;
};

// 余白が無い場合の予備: 期待位置に最も近い枠線の外側の縁を境界にする。
export const chooseLineEdge = (lines, { expected, side, tolerance }) => {
  const candidates = lines
    .map((line) => {
      const edge = side === 'start' ? line.start : line.end;
      return { boundary: side === 'start' ? line.start - 1 : line.end + 1, distance: Math.abs(edge - expected) };
    })
    .filter((candidate) => candidate.distance <= tolerance);
  if (candidates.length === 0) return null;
  candidates.sort((left, right) => left.distance - right.distance);
  return candidates[0].boundary;
};

const snapAxis = ({ ink, dark }, expectedStart, expectedSize, options) => {
  const tolerance = expectedSize * options.searchToleranceRatio;
  const pad = Math.max(1, Math.round(expectedSize * options.paddingRatio));
  const minGutter = Math.max(2, Math.round(expectedSize * options.minGutterRatio));
  const maxThickness = Math.max(2, Math.round(expectedSize * options.maxLineThicknessRatio));
  const expectedEnd = expectedStart + expectedSize;

  const runs = findWhiteRuns(ink, { whiteMax: options.whiteMaxCoverage, minLength: minGutter });
  let start = chooseGutterEdge(runs, { expected: expectedStart, side: 'start', tolerance, pad });
  let end = chooseGutterEdge(runs, { expected: expectedEnd, side: 'end', tolerance, pad });

  if (start === null || end === null) {
    const lines = detectFrameLines(dark, { minCoverage: options.minLineCoverage, maxThickness });
    if (start === null) start = chooseLineEdge(lines, { expected: expectedStart, side: 'start', tolerance });
    if (end === null) end = chooseLineEdge(lines, { expected: expectedEnd, side: 'end', tolerance });
  }

  let resolvedStart = start === null ? expectedStart : Math.max(0, start);
  let resolvedEnd = end === null ? expectedEnd : Math.min(ink.length, end + 1);
  const size = resolvedEnd - resolvedStart;
  const withinTolerance = size > 0 && Math.abs(size - expectedSize) <= expectedSize * options.sizeToleranceRatio;
  if (!withinTolerance) {
    resolvedStart = expectedStart;
    resolvedEnd = expectedEnd;
  }
  return {
    start: resolvedStart,
    size: resolvedEnd - resolvedStart,
    snappedStart: withinTolerance && start !== null,
    snappedEnd: withinTolerance && end !== null
  };
};

// 探索窓内のプロファイルと期待枠 (窓内 px 座標) から、余白/枠線にスナップした矩形を返す。
export const snapRectToFrame = ({ profiles, expected, options = {} }) => {
  const resolved = { ...DEFAULT_FRAME_SNAP_OPTIONS, ...options };
  const vertical = snapAxis({ ink: profiles.rowsInk, dark: profiles.rowsDark }, expected.y, expected.height, resolved);
  const horizontal = snapAxis({ ink: profiles.colsInk, dark: profiles.colsDark }, expected.x, expected.width, resolved);
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
