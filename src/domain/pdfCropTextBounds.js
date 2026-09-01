import {
  getPdfCropRectFromGrid,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode
} from './pdfCropImport.js';

// --- 誌面の「目印」から切り抜き枠を決める ---
// 各コマには左上に黒四角のコマ番号 (例: [2])、その対角の右下にメーカー名が入る。
// どちらも DTP 上でコマ枠を基準に置かれるため、文字レイヤーの座標を読めばコマの実位置が分かる。
// グリッド推定より精度が高いので、目印が見つかったコマではこちらを出発点にする。
//
// 流れ:
//   1. 期待枠の左上付近で「コマ番号と一致する数字」を探す (= 左上アンカー)
//   2. 期待枠の中でコードラベル (261-Exxxx) を探す (= 左端の裏付け)
//   3. アンカーからコマ 1 つ分の範囲にある文字を集め、その外接矩形 + わずかな余白を切り抜き枠にする
//      隣のコマの目印が見つかっていれば、その手前を収集範囲の限界にする
//   4. 期待枠から大きく外れる結果は捨てる (グリッド推定へフォールバック)

export const DEFAULT_PDF_TEXT_BOUNDS_OPTIONS = Object.freeze({
  ascentRatio: 0.9,        // baseline から文字上端まで (文字高さ比)
  descentRatio: 0.22,      // baseline から文字下端まで
  cornerReachRatio: 0.4,   // 左上アンカーを探す範囲 (コマ幅/高さ比)
  slackRatio: 0.05,        // 文字を集めるときの許容はみ出し (コマ幅/高さ比)
  marginXRatio: 0.03,      // 文字の外側に付ける余白 (コマ幅比)
  marginYRatio: 0.026,     // 文字の外側に付ける余白 (コマ高さ比)
  minSizeRatio: 0.7,       // 期待サイズに対する許容下限
  maxSizeRatio: 1.15,      // 期待サイズに対する許容上限
  maxShiftRatio: 0.3       // 期待位置からのずれ許容 (辺長比)
});

const toNumber = (value) => (Number.isFinite(Number(value)) ? Number(value) : 0);

// 文字アイテム (x = 左端, y = ベースライン) から外接ボックスを求める
export const getPdfTextItemBox = (item, options = DEFAULT_PDF_TEXT_BOUNDS_OPTIONS) => {
  const x = toNumber(item?.x);
  const y = toNumber(item?.y);
  const width = Math.max(0, toNumber(item?.width));
  const height = Math.max(0, toNumber(item?.height));
  return {
    left: x,
    right: x + width,
    top: y - height * options.ascentRatio,
    bottom: y + height * options.descentRatio
  };
};

const normalizeLabel = (value) => String(value ?? '').normalize('NFKC').replace(/\s+/g, '');
// CSV のコマ番号 / 追番は未取得のとき 0 になるため、0 は目印として使わない
const isFrameLabel = (value) => /^[1-9]\d{0,2}$/.test(value);

// 期待枠の左上付近にある、コマ番号と一致する数字を返す
export const findPdfCropCornerAnchor = (boxes, { labels = [], expected, options = DEFAULT_PDF_TEXT_BOUNDS_OPTIONS } = {}) => {
  const reachX = expected.width * options.cornerReachRatio;
  const reachY = expected.height * options.cornerReachRatio;
  // コマ番号 → 追番の順に試す (先に見つかった方を採用し、別コマの番号との取り違えを避ける)
  for (const label of labels.map(normalizeLabel).filter(isFrameLabel)) {
    let best = null;
    boxes.forEach((entry) => {
      if (entry.label !== label) return;
      const dx = entry.box.left - expected.x;
      const dy = entry.box.top - expected.y;
      if (Math.abs(dx) > reachX || Math.abs(dy) > reachY) return;
      const distance = Math.hypot(dx / expected.width, dy / expected.height);
      if (!best || distance < best.distance) best = { ...entry, distance };
    });
    if (best) return best;
  }
  return null;
};

// 期待枠の中にあるコードラベルを返す (左下に最も近いもの)
export const findPdfCropCodeAnchor = (boxes, { code, expected, options = DEFAULT_PDF_TEXT_BOUNDS_OPTIONS } = {}) => {
  if (!code) return null;
  const slackX = expected.width * options.slackRatio;
  const slackY = expected.height * options.slackRatio;
  let best = null;
  boxes.forEach((entry) => {
    if (entry.code !== code) return;
    const { box } = entry;
    if (box.left < expected.x - slackX || box.left > expected.x + expected.width + slackX) return;
    if (box.top < expected.y - slackY || box.top > expected.y + expected.height + slackY) return;
    const distance = Math.hypot(
      (box.left - expected.x) / expected.width,
      (box.bottom - (expected.y + expected.height)) / expected.height
    );
    if (!best || distance < best.distance) best = { ...entry, distance };
  });
  return best;
};

const isPlausibleRect = (rect, expected, options) => (
  rect.width > 0
  && rect.height > 0
  && rect.width >= expected.width * options.minSizeRatio
  && rect.width <= expected.width * options.maxSizeRatio
  && rect.height >= expected.height * options.minSizeRatio
  && rect.height <= expected.height * options.maxSizeRatio
  && Math.abs(rect.x - expected.x) <= expected.width * options.maxShiftRatio
  && Math.abs(rect.y - expected.y) <= expected.height * options.maxShiftRatio
);

// 隣のコマの左上アンカーが見つかっていれば、その手前を文字収集範囲の限界にする
const limitByNeighbours = (target, targets, base, options) => {
  const { expected } = target;
  const slackX = expected.width * options.slackRatio;
  const slackY = expected.height * options.slackRatio;
  let right = base.left + expected.width + slackX;
  let bottom = base.top + expected.height + slackY;
  targets.forEach((other) => {
    if (other === target || !other.corner) return;
    const { box } = other.corner;
    const overlapsRow = box.top < base.top + expected.height * 0.9 && box.bottom > base.top;
    const overlapsColumn = box.left < base.left + expected.width * 0.9 && box.right > base.left;
    if (overlapsRow && box.left > base.left + expected.width * 0.5) right = Math.min(right, box.left - slackX);
    if (overlapsColumn && box.top > base.top + expected.height * 0.5) bottom = Math.min(bottom, box.top - slackY);
  });
  return { right, bottom };
};

const buildRectFromAnchors = (target, targets, boxes, options) => {
  const { expected, corner, codeAnchor } = target;
  if (!corner && !codeAnchor) return null;

  const anchorLefts = [corner?.box.left, codeAnchor?.box.left].filter(Number.isFinite);
  const base = { left: Math.min(...anchorLefts), top: corner ? corner.box.top : expected.y };
  const slackX = expected.width * options.slackRatio;
  const slackY = expected.height * options.slackRatio;
  const limit = limitByNeighbours(target, targets, base, options);

  const inside = boxes.filter(({ box }) => (
    box.left >= base.left - slackX
    && box.right <= limit.right
    && box.top >= base.top - slackY
    && box.bottom <= limit.bottom
  ));
  if (inside.length === 0) return null;

  const contentLeft = Math.min(base.left, ...inside.map(({ box }) => box.left));
  const contentTop = corner ? base.top : Math.min(...inside.map(({ box }) => box.top));
  const contentRight = Math.max(...inside.map(({ box }) => box.right));
  const contentBottom = Math.max(...inside.map(({ box }) => box.bottom));

  const marginX = expected.width * options.marginXRatio;
  const marginY = expected.height * options.marginYRatio;
  const rect = {
    x: contentLeft - marginX,
    y: contentTop - marginY,
    width: contentRight - contentLeft + marginX * 2,
    height: contentBottom - contentTop + marginY * 2
  };
  return isPlausibleRect(rect, expected, options) ? rect : null;
};

// ページ内の各コマについて、目印から求めた切り抜き枠を Map<rowId, rect> で返す。
// 目印が足りない / 結果が期待枠から大きく外れるコマは含めない (呼び出し側がグリッド推定へフォールバックする)。
export const resolvePdfCropTextRects = ({ rows = [], textItems = [], grid, options = {} } = {}) => {
  const resolved = { ...DEFAULT_PDF_TEXT_BOUNDS_OPTIONS, ...options };
  const rects = new Map();
  if (!grid || !Array.isArray(textItems) || textItems.length === 0) return rects;

  const boxes = textItems
    .filter((item) => item && String(item.text || '').trim())
    .map((item) => ({
      box: getPdfTextItemBox(item, resolved),
      label: normalizeLabel(item.text),
      code: normalizePdfCropCode(item.text)
    }));

  const targets = rows.filter(isPdfCropRowInsideGrid).map((row) => {
    const expected = getPdfCropRectFromGrid(row, grid);
    return {
      row,
      expected,
      corner: findPdfCropCornerAnchor(boxes, { labels: [row.frameNumber, row.order], expected, options: resolved }),
      codeAnchor: findPdfCropCodeAnchor(boxes, { code: normalizePdfCropCode(row.code), expected, options: resolved })
    };
  });

  targets.forEach((target) => {
    const rect = buildRectFromAnchors(target, targets, boxes, resolved);
    if (rect) rects.set(target.row.id, rect);
  });
  return rects;
};

// ピクセルから決まった枠と目印の枠を合わせる。
// 罫線や余白を検出できた辺はコマの区切りそのものなのでその位置を採用し、
// 検出できなかった辺だけ目印の枠まで広げて文字が欠けないようにする。
export const mergePdfCropRects = (detected, textRect, detectedEdges = {}) => {
  if (!detected) return textRect || null;
  if (!textRect) return detected;
  const left = detectedEdges.left ? detected.x : Math.min(detected.x, textRect.x);
  const top = detectedEdges.top ? detected.y : Math.min(detected.y, textRect.y);
  const right = detectedEdges.right
    ? detected.x + detected.width
    : Math.max(detected.x + detected.width, textRect.x + textRect.width);
  const bottom = detectedEdges.bottom
    ? detected.y + detected.height
    : Math.max(detected.y + detected.height, textRect.y + textRect.height);
  return { x: left, y: top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
};
