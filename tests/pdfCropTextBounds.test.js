import test from 'node:test';
import assert from 'node:assert/strict';

import {
  DEFAULT_PDF_TEXT_BOUNDS_OPTIONS as OPTIONS,
  findPdfCropCornerAnchor,
  getPdfTextItemBox,
  resolvePdfCropTextRects,
  unionPdfCropRects
} from '../src/domain/pdfCropTextBounds.js';
import { getPdfCropRectFromGrid } from '../src/domain/pdfCropImport.js';

const GRID = { left: 0.1, top: 0.06, cellWidth: 0.2, cellHeight: 0.2 };
const TEXT_HEIGHT = 0.008;

// 文字アイテムを「左端・上端」で指定する (実データの y はベースライン)
const item = (text, left, top, width, height = TEXT_HEIGHT) => ({
  text,
  x: left,
  y: top + height * OPTIONS.ascentRatio,
  width,
  height
});
const boxBottom = (top, height = TEXT_HEIGHT) => top + height * (OPTIONS.ascentRatio + OPTIONS.descentRatio);

const makeRow = (overrides) => ({
  id: 'row-1',
  code: 'E1894',
  frameNumber: 2,
  order: 2,
  xPos: 1,
  yPos: 1,
  rowSpan: 1,
  colSpan: 1,
  layoutStatus: 'ready',
  ...overrides
});

// コマ 1 (左上) とその真下のコマ 2。誌面と同じく左上に番号、左下にコード、右下にメーカー名。
const PANEL_ONE = makeRow({ id: 'p1', code: 'E1894', frameNumber: 2, order: 2, yPos: 1 });
const PANEL_TWO = makeRow({ id: 'p2', code: 'E1900', frameNumber: 6, order: 6, yPos: 2 });
const PAGE_ITEMS = [
  item('2', 0.108, 0.066, 0.008),
  item('やさしくラクケアシリーズ まるで果物のようなゼリー', 0.12, 0.0665, 0.16),
  item('261-E1894', 0.108, 0.2, 0.05),
  item('¥2,139', 0.25, 0.222, 0.042),
  item('ハウスギャバン(株)', 0.24, 0.24, 0.052),
  item('6', 0.108, 0.266, 0.008),
  item('261-E1900', 0.108, 0.4, 0.05),
  item('別メーカー(株)', 0.242, 0.44, 0.05)
];

const near = (actual, expected, epsilon = 1e-9) => assert.ok(
  Math.abs(actual - expected) < epsilon,
  `${actual} is not close to ${expected}`
);

test('getPdfTextItemBox expands the baseline into an ascent/descent box', () => {
  const box = getPdfTextItemBox(item('あ', 0.2, 0.5, 0.03));
  near(box.left, 0.2);
  near(box.right, 0.23);
  near(box.top, 0.5);
  near(box.bottom, boxBottom(0.5));
});

test('resolvePdfCropTextRects crops from the corner number to the maker name with a small margin', () => {
  const rects = resolvePdfCropTextRects({ rows: [PANEL_ONE, PANEL_TWO], textItems: PAGE_ITEMS, grid: GRID });
  const rect = rects.get('p1');
  assert.ok(rect, 'コマ 1 の枠が求まること');

  const expected = getPdfCropRectFromGrid(PANEL_ONE, GRID);
  const marginX = expected.width * OPTIONS.marginXRatio;
  const marginY = expected.height * OPTIONS.marginYRatio;
  // 左上 = 番号ラベルの左上、右下 = メーカー名の右下 (どちらも余白ぶん外側)
  near(rect.x, 0.108 - marginX);
  near(rect.y, 0.066 - marginY);
  near(rect.x + rect.width, 0.24 + 0.052 + marginX);
  near(rect.y + rect.height, boxBottom(0.24) + marginY);
  // 下のコマの番号ラベル / メーカー名は入り込まない
  assert.ok(rect.y + rect.height < 0.266);
});

test('resolvePdfCropTextRects keeps panels independent on the same page', () => {
  const rects = resolvePdfCropTextRects({ rows: [PANEL_ONE, PANEL_TWO], textItems: PAGE_ITEMS, grid: GRID });
  const second = rects.get('p2');
  assert.ok(second, 'コマ 2 の枠が求まること');
  const marginY = GRID.cellHeight * OPTIONS.marginYRatio;
  near(second.y, 0.266 - marginY);
  near(second.y + second.height, boxBottom(0.44) + marginY);
});

test('resolvePdfCropTextRects falls back to the code label when the corner number is missing', () => {
  const withoutCorner = PAGE_ITEMS.filter((entry) => entry.text !== '2');
  const rects = resolvePdfCropTextRects({ rows: [PANEL_ONE, PANEL_TWO], textItems: withoutCorner, grid: GRID });
  const rect = rects.get('p1');
  assert.ok(rect, 'コードラベルだけでも枠が求まること');
  const marginY = GRID.cellHeight * OPTIONS.marginYRatio;
  // 上端は番号ラベルではなく、収集した文字の最上端 (見出し) から決まる
  near(rect.y, 0.0665 - marginY);
});

test('resolvePdfCropTextRects ignores a digit that is far from the panel corner', () => {
  // コマ左上の「2」を、コマ中央より右下にある「2」(単位や価格の一部) に置き換える
  const strayDigit = PAGE_ITEMS.map((entry) => (
    entry.text === '2' && entry.x < 0.15 ? item('2', 0.24, 0.2, 0.008) : entry
  ));
  const rect = resolvePdfCropTextRects({ rows: [PANEL_ONE, PANEL_TWO], textItems: strayDigit, grid: GRID }).get('p1');
  assert.ok(rect);
  // 左上アンカーとして採用されていないので、上端は最上端の文字 (見出し) 基準になる
  near(rect.y, 0.0665 - GRID.cellHeight * OPTIONS.marginYRatio);
});

test('resolvePdfCropTextRects drops results that are implausibly small or shifted', () => {
  // 目印だけで中身が無い → 期待サイズの 70% を下回るので採用しない
  const sparse = [item('2', 0.108, 0.066, 0.008), item('261-E1894', 0.108, 0.09, 0.05)];
  assert.equal(resolvePdfCropTextRects({ rows: [PANEL_ONE], textItems: sparse, grid: GRID }).size, 0);
  // 目印がまったく無いページは空 (グリッド推定にフォールバック)
  assert.equal(resolvePdfCropTextRects({ rows: [PANEL_ONE], textItems: [item('無関係', 0.5, 0.5, 0.1)], grid: GRID }).size, 0);
  assert.equal(resolvePdfCropTextRects({ rows: [PANEL_ONE], textItems: [], grid: GRID }).size, 0);
});

test('findPdfCropCornerAnchor takes a real frame number at the corner only', () => {
  const expected = getPdfCropRectFromGrid(PANEL_ONE, GRID);
  const corner = { left: 0.108, top: 0.066, right: 0.116, bottom: 0.076 };
  const middle = { left: 0.24, top: 0.2, right: 0.248, bottom: 0.21 };
  const boxes = [
    { label: '0', code: '', box: corner },
    { label: '2', code: '', box: corner },
    { label: '6', code: '', box: middle }
  ];
  // CSV にコマ番号/追番が無い行は 0 になる。0 は目印にしない
  assert.equal(findPdfCropCornerAnchor(boxes, { labels: [0, 0], expected }), null);
  assert.equal(findPdfCropCornerAnchor(boxes, { labels: [2], expected }).box, corner);
  // コマ番号で見つからなければ追番で探す
  assert.equal(findPdfCropCornerAnchor(boxes, { labels: [99, 2], expected }).box, corner);
  // 角から離れた同じ数字は使わない
  assert.equal(findPdfCropCornerAnchor(boxes, { labels: [6], expected }), null);
});

test('unionPdfCropRects keeps both rectangles inside and tolerates a missing side', () => {
  const text = { x: 0.1, y: 0.2, width: 0.2, height: 0.2 };
  const snapped = { x: 0.09, y: 0.21, width: 0.22, height: 0.2 };
  assert.deepEqual(unionPdfCropRects(snapped, text), {
    x: 0.09,
    y: 0.2,
    width: 0.22,
    height: 0.21000000000000002
  });
  assert.equal(unionPdfCropRects(snapped, null), snapped);
  assert.equal(unionPdfCropRects(null, text), text);
});
