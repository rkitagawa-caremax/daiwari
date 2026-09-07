import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildEdgeCatalogProducts,
  buildCatalogChangeSet,
  compareCatalogSnapshots,
  cosineSimilarity,
  parseCatalogSnapshotCsv,
  parseCatalogSnapshotRecords,
  parseCsvRecords,
  rankEdgeCatalogProducts,
  summarizeCatalogDiff
} from '../src/domain/edgeAiCatalog.js';

test('buildEdgeCatalogProducts combines saved text, placement and sales data', () => {
  const images = [{
    id: 'image-1',
    data: 'data:image/jpeg;base64,abc',
    code: 'e001',
    productName: '軽量歩行器',
    sourceText: '折りたたみできます',
    sizeType: '1/4',
    catalogTextData: {
      itemNumber: 'KW-1',
      specifications: ['重量5kg'],
      handlingMarkers: ['●屋外'],
      compositionDetails: [],
      materialDetails: [],
      availabilityLabels: [],
      itemNumberCandidates: [],
      catchCopyCandidates: []
    }
  }];
  const sheets = [{ id: 'sheet-1', genre: 'walk', panels: [{ imageId: 'image-1', code: 'E001', rowSpan: 2, colSpan: 2, sizeType: '1/4（4コマ）' }] }];
  const products = buildEdgeCatalogProducts({
    images,
    sheets,
    genres: [{ id: 'walk', label: '歩行関連' }],
    salesData: { E001: [{
      name: '軽量歩行器',
      spec: '青',
      count: 12,
      salesAmount: 120000,
      grossProfitAmount: 36000,
      monthlySales: [5, 7],
      monthlyLabels: ['4月', '5月']
    }] }
  });
  assert.equal(products.length, 1);
  assert.equal(products[0].code, 'E001');
  assert.equal(products[0].salesCount, 12);
  assert.equal(products[0].salesAmount, 120000);
  assert.equal(products[0].grossProfitAmount, 36000);
  assert.equal(products[0].grossMargin, 0.3);
  assert.equal(products[0].catalogTextCompleteness > 0, true);
  assert.equal(products[0].salesMatched, true);
  assert.deepEqual(products[0].monthlySales, [
    { label: '4月', count: 5 },
    { label: '5月', count: 7 }
  ]);
  assert.equal(products[0].assignments[0].pageNumber, 1);
  assert.equal(products[0].assignments[0].rowSpan, 2);
  assert.equal(products[0].assignments[0].colSpan, 2);
  assert.match(products[0].searchText, /重量5kg/);
});

test('parseCsvRecords parses quoted multiline cells', () => {
  assert.deepEqual(parseCsvRecords('code,text\r\nE001,"1行目\n2行目"'), [
    ['code', 'text'],
    ['E001', '1行目\n2行目']
  ]);
});

test('parseCatalogSnapshotCsv recognizes Japanese catalog headers', () => {
  const parsed = parseCatalogSnapshotCsv('\uFEFF介援隊コード,商品名,税込価格,仕様,デモ機\nE001,車いす,"12,800円",軽量・折りたたみ,あり');
  assert.equal(parsed.items.length, 1);
  assert.deepEqual(parsed.items[0], {
    code: 'E001',
    name: '車いす',
    itemNumber: '',
    catchCopy: '',
    priceIncludingTax: 12800,
    priceExcludingTax: '',
    availability: '',
    lifecycleStatus: '',
    handlingMarkers: [],
    demoStatus: 'D',
    hasDemoMarker: true,
    specifications: ['軽量', '折りたたみ'],
    compositionDetails: [],
    materialDetails: []
  });
});

test('catalog snapshot records detect a heading row and retain only supplied comparison fields', () => {
  const parsed = parseCatalogSnapshotRecords([
    ['価格改定表', '', ''],
    ['更新日', '2026/09/03', ''],
    ['介援隊コード', '税込価格', 'デモ機区分'],
    ['E001', 1540, '(N)']
  ]);
  assert.equal(parsed.headerRowIndex, 2);
  assert.deepEqual(parsed.recognizedFields, ['priceIncludingTax', 'demoStatus']);
  assert.equal(parsed.items[0].priceIncludingTax, 1540);
  assert.equal(parsed.items[0].demoStatus, 'N');
});

test('catalog change set compares uploaded fields only and does not infer missing rows as discontinued', () => {
  const products = [
    { code: 'E001', name: '旧商品名', priceIncludingTax: 1408, priceExtractionConfidence: 'medium', lifecycleStatus: '通常', demoStatus: 'D' },
    { code: 'E002', name: '掲載中', priceIncludingTax: 2200, demoStatus: '' }
  ];
  const changeSet = buildCatalogChangeSet({
    products,
    fileName: '価格表.xlsx',
    sheetName: '改定',
    snapshot: {
      items: [{ code: 'E001', name: '', priceIncludingTax: 1540, lifecycleStatus: '廃盤', demoStatus: 'N' }],
      recognizedFields: ['priceIncludingTax', 'lifecycleStatus', 'demoStatus']
    }
  });
  assert.equal(changeSet.diffs.length, 1);
  assert.equal(changeSet.displayCount, 1);
  assert.deepEqual(changeSet.diffs[0].changes.map((change) => change.key), ['priceIncludingTax', 'lifecycleStatus', 'demoStatus']);
  assert.deepEqual(changeSet.diffs[0].changes.map((change) => [change.before, change.after]), [
    ['1408', '1540'],
    ['通常', '廃盤'],
    ['(D)', '(N)']
  ]);
  assert.equal(changeSet.diffs[0].changes[0].confidence, 'medium');
  assert.equal(changeSet.byCode.E002, undefined);
});

test('compareCatalogSnapshots reports additions removals and field changes', () => {
  const before = [
    { code: 'E001', name: '旧商品', priceIncludingTax: 1000 },
    { code: 'E002', name: '終了品' }
  ];
  const after = [
    { code: 'E001', name: '新商品', priceIncludingTax: 1200 },
    { code: 'E003', name: '追加品' }
  ];
  const diffs = compareCatalogSnapshots(before, after);
  assert.deepEqual(summarizeCatalogDiff(diffs), { added: 1, removed: 1, modified: 1, unchanged: 0 });
  const modified = diffs.find((diff) => diff.code === 'E001');
  assert.deepEqual(modified.changes.map((change) => change.key), ['name', 'priceIncludingTax']);
  assert.equal(modified.changes[1].severity, 'high');
});

test('rankEdgeCatalogProducts gives exact code and matching description priority', () => {
  const products = [
    { id: '1', code: 'E001', name: '標準車いす', searchText: '標準車いす 自走式', salesCount: 0 },
    { id: '2', code: 'E002', name: '歩行器', searchText: '軽量 折りたたみ 歩行器', salesCount: 0 }
  ];
  assert.equal(rankEdgeCatalogProducts(products, 'E002')[0].product.id, '2');
  assert.equal(rankEdgeCatalogProducts(products, '折りたたみ')[0].product.id, '2');
});

test('cosineSimilarity compares normalized directions', () => {
  assert.equal(cosineSimilarity([1, 0], [1, 0]), 1);
  assert.equal(cosineSimilarity([1, 0], [0, 1]), 0);
});
