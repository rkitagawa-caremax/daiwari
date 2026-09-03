import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildEdgeCatalogProducts,
  compareCatalogSnapshots,
  cosineSimilarity,
  parseCatalogSnapshotCsv,
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
  const sheets = [{ id: 'sheet-1', genre: 'walk', panels: [{ imageId: 'image-1', code: 'E001' }] }];
  const products = buildEdgeCatalogProducts({
    images,
    sheets,
    genres: [{ id: 'walk', label: '歩行関連' }],
    salesData: { E001: [{ name: '軽量歩行器', spec: '青', count: 12 }] }
  });
  assert.equal(products.length, 1);
  assert.equal(products[0].code, 'E001');
  assert.equal(products[0].salesCount, 12);
  assert.equal(products[0].assignments[0].pageNumber, 1);
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
    handlingMarkers: [],
    hasDemoMarker: true,
    specifications: ['軽量', '折りたたみ'],
    compositionDetails: [],
    materialDetails: []
  });
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
