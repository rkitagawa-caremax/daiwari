import test from 'node:test';
import assert from 'node:assert/strict';

import {
  applyPdfCropCatalogPageOverrides,
  buildPdfCropBatchPages,
  buildPdfCropPagePlans,
  findPdfCropGridConflicts,
  summarizePdfCropPagePlans,
  DEFAULT_PDF_GRID_BOUNDS,
  extractPdfCatalogTextData,
  extractPdfCatalogDetails,
  extractPdfCropText,
  extractPdfPriceFields,
  extractPdfTextInRect,
  getPdfCropRect,
  inferCatalogStartPage,
  MAX_PDF_CROP_BATCH_PAGES,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode,
  parsePdfCropCsv,
  resolvePdfCropCsvColumns,
  resolvePdfTextExtractionRect,
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
    catalogName: 'やさしくラクケアシリーズ まるで果物のようなゼリー',
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
    catalogName: 9,
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
    catalogName: -1,
    coordinate: 8,
    xPos: 10,
    yPos: 11
  });
});

test('PDF text extraction excludes a page-wide header whose center is outside the panel', () => {
  const rect = { x: 0.1, y: 0.1, width: 0.2, height: 0.2 };
  const result = extractPdfTextInRect([
    { text: 'ページ全体の校正指示', x: 0.05, y: 0.12, width: 0.9, height: 0.01 },
    { text: '商品名', x: 0.12, y: 0.14, width: 0.08, height: 0.01 },
    { text: '261-E1955', x: 0.12, y: 0.25, width: 0.06, height: 0.01 }
  ], rect);

  assert.equal(result.text, '商品名 261-E1955');
});

test('PDF price extraction separates tax-included and tax-excluded values and keeps candidates', () => {
  const result = extractPdfPriceFields(
    '明治メイバランス 261-E1955 各 ¥ 5,832 (税抜¥5,400) 参考価格 ¥6,000',
    'E1955'
  );

  assert.equal(result.priceIncludingTax, 5832);
  assert.equal(result.priceExcludingTax, 5400);
  assert.equal(result.priceExtractionConfidence, 'high');
  assert.deepEqual(result.priceCandidates, [
    { amount: 5832, taxType: 'including', text: '¥5,832' },
    { amount: 5400, taxType: 'excluding', text: '¥5,400' },
    { amount: 6000, taxType: 'including', text: '¥6,000' }
  ]);
});

test('PDF catalog text data uses the CSV catalog name and stores extraction version', () => {
  const result = extractPdfCatalogTextData({
    textItems: [
      { text: '261-E1955', x: 0.12, y: 0.15, width: 0.06, height: 0.01 },
      { text: '¥5,832', x: 0.12, y: 0.2, width: 0.04, height: 0.01 }
    ],
    rect: { x: 0.1, y: 0.1, width: 0.2, height: 0.2 },
    code: '261-E1955',
    catalogName: ' 明治メイバランスブリックゼリー '
  });

  assert.equal(result.catalogCode, 'E1955');
  assert.equal(result.productName, '明治メイバランスブリックゼリー');
  assert.equal(result.productNameSource, 'csv');
  assert.equal(result.sourceTextVersion, 2);
  assert.equal(result.priceIncludingTax, 5832);
});

test('PDF catalog details retain item numbers, handling markers, stock status and bullet specifications', () => {
  const result = extractPdfCatalogDetails(
    '使い捨て防水エプロン。食べこぼしを吸収するロングエプロン。 261-E0503 KN-932 10枚入 ¥1,028 (税抜¥935) ●材質/表面:パルプ、裏面:ポリエチレンラミネート ●成分/パルプ、ポリエチレン ●生産国/日本 ●50 (D) (株)ストリックスデザイン 在庫商品',
    { code: 'E0503', productName: '使い捨て防水エプロン' }
  );

  assert.equal(result.itemNumber, 'KN-932');
  assert.ok(result.itemNumberCandidates.includes('KN-932'));
  assert.equal(result.availability, 'stock');
  assert.deepEqual(result.availabilityLabels, ['在庫商品']);
  assert.ok(result.handlingMarkers.includes('(D)'));
  assert.equal(result.hasDemoMarker, true);
  assert.ok(result.specifications.includes('●材質/表面:パルプ、裏面:ポリエチレンラミネート'));
  assert.deepEqual(result.compositionDetails, ['●成分/パルプ、ポリエチレン']);
  assert.deepEqual(result.materialDetails, ['●材質/表面:パルプ、裏面:ポリエチレンラミネート']);
  assert.equal(result.catchCopy, '食べこぼしを吸収するロングエプロン。');
});

test('PDF catalog details preserve numeric and unhyphenated item-number candidates and direct shipping', () => {
  const numeric = extractPdfCatalogDetails('261-E1911 92084 ¥1,408 (税抜¥1,280) メーカー直送', { code: 'E1911' });
  const alphaNumeric = extractPdfCatalogDetails('261-S1090 SWR142SAL U型シート ¥39,600', { code: 'S1090' });
  const copyAfterPrice = extractPdfCatalogDetails(
    '536-050、536-051 261-S0793 各¥39,600 (税抜¥36,000) これなら邪魔にならない♪最薄クラスの折りたたみ幅15cm。',
    { code: 'S0793' }
  );

  assert.equal(numeric.itemNumber, '92084');
  assert.equal(numeric.availability, 'direct');
  assert.deepEqual(numeric.availabilityLabels, ['直送']);
  assert.ok(alphaNumeric.itemNumberCandidates.includes('SWR142SAL'));
  assert.deepEqual(copyAfterPrice.catchCopyCandidates.slice(0, 2), [
    'これなら邪魔にならない♪',
    '最薄クラスの折りたたみ幅15cm。'
  ]);
});

test('PDF catalog details retain discontinued and limited-stock lifecycle labels', () => {
  assert.equal(extractPdfCatalogDetails('261-E1001 在庫限り').lifecycleStatus, '在庫限り');
  assert.equal(extractPdfCatalogDetails('261-E1002 廃盤').lifecycleStatus, '廃盤');
});

test('automatic text extraction rect is limited by grid and text anchors while manual stays unchanged', () => {
  const cropRect = { x: 0.08, y: 0.04, width: 0.28, height: 0.3 };
  const gridRect = { x: 0.1, y: 0.06, width: 0.2, height: 0.2 };
  const textRect = { x: 0.11, y: 0.07, width: 0.18, height: 0.18 };

  assert.deepEqual(resolvePdfTextExtractionRect({ cropRect, gridRect, textRect }), textRect);
  assert.equal(resolvePdfTextExtractionRect({ cropRect, gridRect, textRect, isManual: true }), cropRect);
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
