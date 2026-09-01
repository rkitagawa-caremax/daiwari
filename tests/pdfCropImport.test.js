import test from 'node:test';
import assert from 'node:assert/strict';

import {
  applyPdfCropCatalogPageOverrides,
  buildPdfCropBatchPages,
  buildPdfCropPagePlans,
  findPdfCropGridConflicts,
  summarizePdfCropPagePlans,
  DEFAULT_PDF_GRID_BOUNDS,
  extractPdfCropText,
  getPdfCropRect,
  inferCatalogStartPage,
  MAX_PDF_CROP_BATCH_PAGES,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode,
  parsePdfCropCsv,
  resolvePdfCropCsvColumns,
  pdfTextItemsContainCode,
  updatePdfCropRowSize
} from '../src/domain/pdfCropImport.js';

const HEADER = 'ジャンル,ページ数,追番,コマ番号,介援隊コード,コマ数,,テキスト情報,座標,コマID,X_POS,Y_POS';

test('normalizePdfCropCode removes catalog prefixes and proof suffixes', () => {
  assert.equal(normalizePdfCropCode('261-E1957'), 'E1957');
  assert.equal(normalizePdfCropCode('010-E1894_'), 'E1894');
  assert.equal(normalizePdfCropCode('ａ１２３４'), 'A1234');
  assert.equal(normalizePdfCropCode('ダミーコマ'), '');
});

test('inferCatalogStartPage reads the P number from a proof PDF filename', () => {
  assert.equal(inferCatalogStartPage('P010_介援隊vol.26.pdf'), 10);
  assert.equal(inferCatalogStartPage('catalog.pdf'), 1);
});

test('buildPdfCropBatchPages maps selected PDFs to consecutive catalog pages', () => {
  const pages = buildPdfCropBatchPages([
    { file: { name: 'P010_proof.pdf' }, numPages: 1 },
    { file: { name: 'P020_21.pdf' }, numPages: 2 }
  ]);
  assert.equal(MAX_PDF_CROP_BATCH_PAGES, 10);
  assert.deepEqual(pages.map(({ filename, pdfPageNumber, catalogPage }) => ({
    filename,
    pdfPageNumber,
    catalogPage
  })), [
    { filename: 'P010_proof.pdf', pdfPageNumber: 1, catalogPage: 10 },
    { filename: 'P020_21.pdf', pdfPageNumber: 1, catalogPage: 20 },
    { filename: 'P020_21.pdf', pdfPageNumber: 2, catalogPage: 21 }
  ]);
});

test('parsePdfCropCsv reads code, size, page and explicit grid coordinates', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,261-E1957,1/8 横（2コマ）,,,X1Y1,PID-1,1,1',
    '食事関連,10,2,2,E1894,1/8 横（2コマ）,,,X3Y1,PID-2,3,1',
    '食事関連,10,3,3,ダミーコマ,1/16（1コマ）,タイトル,,X1Y2,PID-3,1,2',
    '食事関連,10,4,4,E9999,1/16（1コマ）,,説明文,X2Y2,PID-4,2,2'
  ].join('\n');

  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(issues.length, 0);
  assert.equal(rows.length, 2, 'dummy and text rows are not image crop targets');
  assert.deepEqual(rows[0], {
    id: '10-E1957-2',
    csvRow: 2,
    pageNumber: 10,
    order: 1,
    frameNumber: 1,
    rawCode: '261-E1957',
    code: 'E1957',
    filename: 'E1957.jpg',
    sizeType: '1/8 横（2コマ）',
    rowSpan: 1,
    colSpan: 2,
    xPos: 1,
    yPos: 1,
    positionSource: 'csv',
    layoutStatus: 'ready'
  });
});

test('parsePdfCropCsv estimates missing positions in frame order', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,E1957,1/8 横（2コマ）,,,,,,',
    '食事関連,10,2,2,E1894,1/8 横（2コマ）,,,,,,'
  ].join('\n');
  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(issues.length, 0);
  assert.deepEqual(rows.map(({ xPos, yPos, positionSource }) => ({ xPos, yPos, positionSource })), [
    { xPos: 1, yPos: 1, positionSource: 'estimated' },
    { xPos: 3, yPos: 1, positionSource: 'estimated' }
  ]);
});

test('parsePdfCropCsv reports overlapping explicit positions', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,E1957,1/8 横（2コマ）,,,X1Y1,,,',
    '食事関連,10,2,2,E1894,1/16（1コマ）,,,X2Y1,,,'
  ].join('\n');
  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(rows[1].layoutStatus, 'conflict');
  assert.equal(issues[0].type, 'layout-conflict');
});

test('getPdfCropRect converts 4x4 coordinates and spans to normalized bounds', () => {
  const row = { xPos: 3, yPos: 2, rowSpan: 2, colSpan: 2, layoutStatus: 'ready' };
  const rect = getPdfCropRect(row, { left: 10, top: 6, right: 10, bottom: 10 });
  assert.deepEqual(rect, { x: 0.5, y: 0.27, width: 0.4, height: 0.42 });
  assert.equal(isPdfCropRowInsideGrid(row), true);
  assert.equal(isPdfCropRowInsideGrid({ ...row, xPos: 4 }), false);
});

test('row size edits update spans and text verification is limited to its crop', () => {
  const base = { code: 'E1957', xPos: 1, yPos: 1, rowSpan: 1, colSpan: 1, layoutStatus: 'ready' };
  const resized = updatePdfCropRowSize(base, '1/8 横（2コマ）');
  assert.equal(resized.colSpan, 2);
  assert.equal(pdfTextItemsContainCode([
    { text: '261-', x: 0.11, y: 0.2 },
    { text: 'E1957', x: 0.2, y: 0.2 },
    { text: 'E1957', x: 0.8, y: 0.8 }
  ], resized, DEFAULT_PDF_GRID_BOUNDS), true);
  assert.equal(pdfTextItemsContainCode([{ text: 'E1957', x: 0.8, y: 0.8 }], resized), false);
});

test('extractPdfCropText keeps only text inside the crop and compacts it for storage', () => {
  const row = { code: 'E1957', xPos: 1, yPos: 1, rowSpan: 1, colSpan: 2, layoutStatus: 'ready' };
  const result = extractPdfCropText([
    { text: 'アイソカル', x: 0.12, y: 0.12 },
    { text: ' ゼリー ', x: 0.2, y: 0.13 },
    { text: '261-E1957', x: 0.3, y: 0.2 },
    { text: '別コマ', x: 0.8, y: 0.8 }
  ], row);
  assert.equal(result.text, 'アイソカルゼリー 261-E1957');
  assert.equal(result.truncated, false);
});

test('extractPdfCropText caps long text and records truncation', () => {
  const row = { code: 'E1957', xPos: 1, yPos: 1, rowSpan: 1, colSpan: 2, layoutStatus: 'ready' };
  assert.deepEqual(extractPdfCropText([{ text: '123456789', x: 0.2, y: 0.2 }], row, DEFAULT_PDF_GRID_BOUNDS, 5), {
    text: '12345',
    truncated: true
  });
});

test('applyPdfCropCatalogPageOverrides replaces only valid overrides', () => {
  const pages = buildPdfCropBatchPages([{ file: { name: 'P010.pdf' }, numPages: 1 }, { file: { name: 'scan.pdf' }, numPages: 1 }]);
  const overridden = applyPdfCropCatalogPageOverrides(pages, { '2-1': 12, '1-1': 'abc' });
  assert.deepEqual(overridden.map((page) => page.catalogPage), [10, 12]);
  assert.equal(overridden[0], pages[0], 'pages without a valid override keep their identity');
});

const PLAN_ROWS = () => parsePdfCropCsv([
  HEADER,
  '食事関連,10,1,1,E1001,1/8 横（2コマ）,,,X1Y1,,1,1',
  '食事関連,10,2,2,E1002,1/8 横（2コマ）,,,X3Y1,,3,1',
  '食事関連,11,1,1,E1002,1/16（1コマ）,,,X1Y1,,1,1',
  '食事関連,11,2,2,E1003,1/16（1コマ）,,,X2Y1,,2,1'
].join('\n')).rows;

const PLAN_PAGES = () => buildPdfCropBatchPages([
  { file: { name: 'P010.pdf' }, numPages: 1 },
  { file: { name: 'P011.pdf' }, numPages: 1 },
  { file: { name: 'P011_copy.pdf' }, numPages: 1 },
  { file: { name: 'P099.pdf' }, numPages: 1 }
]);

test('buildPdfCropPagePlans keeps existing codes by default, dedupes within the batch and flags duplicate targets', () => {
  const plans = buildPdfCropPagePlans({ batchPages: PLAN_PAGES(), rows: PLAN_ROWS(), existingCodes: new Set(['E1001']) });

  assert.deepEqual(plans.map((plan) => plan.importRows.map((row) => row.code)), [
    ['E1001', 'E1002'], // 既存コード E1001 も既定では保存する (追加登録)
    ['E1003'],          // E1002 は P.10 で保存済みなので除外
    [],                 // 同じ P.11 の重複 PDF: コードは全て登場済み
    []                  // CSV に P.99 がない
  ]);
  assert.equal(plans[0].existingCount, 1);
  assert.equal(plans[1].duplicateCodeCount, 1);
  assert.equal(plans[1].isDuplicateCatalogPage, true);
  assert.equal(plans[3].hasCsvRows, false);

  assert.deepEqual(summarizePdfCropPagePlans(plans), {
    pageCount: 4,
    targetCount: 6,
    importCount: 3,
    skippedCount: 3,
    existingCount: 1,
    unplaceableCount: 0,
    duplicateCodeCount: 3,
    conflictCount: 0,
    pagesWithoutCsv: 1,
    duplicateCatalogPages: 2
  });
});

test('buildPdfCropPagePlans can skip existing codes on request', () => {
  const plans = buildPdfCropPagePlans({
    batchPages: PLAN_PAGES(),
    rows: PLAN_ROWS(),
    existingCodes: new Set(['E1001']),
    skipExistingCodes: true
  });
  assert.deepEqual(plans[0].importRows.map((row) => row.code), ['E1002']);
  assert.equal(plans[0].existingCount, 1);
});

test('buildPdfCropPagePlans still crops overlapping rows and only drops rows without a position', () => {
  const rows = parsePdfCropCsv([
    HEADER,
    '食事関連,10,1,1,E1957,1/8 横（2コマ）,,,X1Y1,,,',
    '食事関連,10,2,2,E1894,1/16（1コマ）,,,X2Y1,,,'
  ].join('\n')).rows;
  rows.push({ ...rows[1], id: 'nopos', code: 'E0000', xPos: null, yPos: null, layoutStatus: 'unresolved' });
  const [plan] = buildPdfCropPagePlans({ batchPages: buildPdfCropBatchPages([{ file: { name: 'P010.pdf' }, numPages: 1 }]), rows });
  assert.equal(rows[1].layoutStatus, 'conflict');
  assert.deepEqual(plan.importRows.map((row) => row.code), ['E1957', 'E1894']);
  assert.equal(plan.conflictIds.size, 3, 'overlapping pair + position-less row are reported as warnings');
  assert.ok(plan.conflictIds.has(rows[0].id) && plan.conflictIds.has(rows[1].id));
  assert.equal(plan.unplaceableCount, 1);
});

test('findPdfCropGridConflicts marks overlapping and out-of-grid rows', () => {
  const conflicts = findPdfCropGridConflicts([
    { id: 'a', xPos: 1, yPos: 1, rowSpan: 1, colSpan: 2 },
    { id: 'b', xPos: 2, yPos: 1, rowSpan: 1, colSpan: 1 },
    { id: 'c', xPos: 4, yPos: 4, rowSpan: 1, colSpan: 2 },
    { id: 'd', xPos: 3, yPos: 3, rowSpan: 1, colSpan: 1 }
  ]);
  assert.deepEqual([...conflicts].sort(), ['a', 'b', 'c']);
});

// 台割の別出力フォーマット (コマID/掲載ブロック/ページ番号/X_POS/Y_POS/コマ番号/コマ種別/コマサイズ/介援隊コード/掲載名)
const BLOCK_HEADER = 'コマID,掲載ブロック,ページ番号,X_POS,Y_POS,コマ番号,コマ種別,コマサイズ,介援隊コード,掲載名';

test('parsePdfCropCsv reads the block-layout CSV where the code is in column I', () => {
  const content = [
    BLOCK_HEADER,
    '0064872,食事関連,10,1,1,1,商品,1/8 横（2コマ）,E1957,エンジョイ小さなカップゼリー150',
    '0059340,食事関連,10,3,1,2,商品,1/8 横（2コマ）,E1894,やさしくラクケアシリーズ まるで果物のようなゼリー',
    '0061379,食事関連,10,1,2,3,商品,1/4 横（4コマ）,E1723,アイソカルゼリーハイカロリー',
    '0065422,ネジ,10,1,4,4,タイトル,1/8 横（2コマ）,E9999,タイトル',
    '0065421,ネジ,10,3,4,5,,1/8 横（2コマ）,,'
  ].join('\n');

  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(issues.length, 0);
  assert.deepEqual(rows.map((row) => row.code), ['E1957', 'E1894', 'E1723'], 'タイトル行と空コード行は対象外');
  assert.deepEqual(rows[1], {
    id: '10-E1894-3',
    csvRow: 3,
    pageNumber: 10,
    order: 0,
    frameNumber: 2,
    rawCode: 'E1894',
    code: 'E1894',
    filename: 'E1894.jpg',
    sizeType: '1/8 横（2コマ）',
    rowSpan: 1,
    colSpan: 2,
    xPos: 3,
    yPos: 1,
    positionSource: 'csv',
    layoutStatus: 'ready'
  });
  assert.deepEqual(
    rows[2],
    { ...rows[2], xPos: 1, yPos: 2, rowSpan: 1, colSpan: 4, frameNumber: 3 }
  );
});

test('resolvePdfCropCsvColumns marks absent columns instead of guessing a position', () => {
  const blockColumns = resolvePdfCropCsvColumns(BLOCK_HEADER.split(','));
  assert.deepEqual(blockColumns, {
    pageNumber: 2,
    order: -1,
    frameNumber: 5,
    code: 8,
    sizeType: 7,
    kind: 6,
    text: -1,
    coordinate: -1,
    xPos: 3,
    yPos: 4
  });

  // 旧フォーマットは見出しどおりに読める
  const legacyColumns = resolvePdfCropCsvColumns(HEADER.split(','));
  assert.deepEqual(
    { code: legacyColumns.code, sizeType: legacyColumns.sizeType, text: legacyColumns.text, xPos: legacyColumns.xPos },
    { code: 4, sizeType: 5, text: 7, xPos: 10 }
  );

  // 見出しが読めない CSV は旧フォーマットの列位置にフォールバックする
  assert.deepEqual(resolvePdfCropCsvColumns(['a', 'b', 'c']), {
    pageNumber: 1,
    order: 2,
    frameNumber: 3,
    code: 4,
    sizeType: 5,
    kind: -1,
    text: 7,
    coordinate: 8,
    xPos: 10,
    yPos: 11
  });
});

test('parsePdfCropCsv leaves blank coordinates unplaced when the CSV positions the rest of the page', () => {
  // 一部の行だけ座標が空欄 = グリッド外のコマ。追番順に押し込まず、除外対象にする
  const mixed = [
    BLOCK_HEADER,
    '0000001,災害,278,1,1,1,商品,1/16（1コマ）,O0941,アルファ米',
    '0000002,災害,278,,,2,商品,1/16（1コマ）,O1183,8年保存 大判ウェット'
  ].join('\n');
  const { rows, issues } = parsePdfCropCsv(mixed);
  assert.equal(rows.length, 2);
  assert.deepEqual(rows.map(({ code, layoutStatus }) => ({ code, layoutStatus })), [
    { code: 'O0941', layoutStatus: 'ready' },
    { code: 'O1183', layoutStatus: 'unresolved' }
  ]);
  assert.deepEqual(issues.map((issue) => issue.type), ['layout-unresolved']);

  // 部品表のように 4x4 に収まらないページも推測しない
  const partsList = [
    BLOCK_HEADER,
    ...Array.from({ length: 20 }, (_, index) => (
      `000${index},歩行,158,,,${index + 1},商品,1/16（1コマ）,W${1000 + index},交換ゴム`
    ))
  ].join('\n');
  const partsPage = parsePdfCropCsv(partsList);
  assert.equal(partsPage.rows.filter((row) => row.layoutStatus === 'ready').length, 0);
  assert.equal(partsPage.issues.length, 20);
});
