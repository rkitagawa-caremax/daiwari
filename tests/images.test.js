import test from 'node:test';
import assert from 'node:assert/strict';

import {
  isImageWorkedByUser,
  isSameStockImageList,
  normalizeCloudImageDocuments,
  normalizeStockImageEntry,
  normalizeStockImages
} from '../src/domain/images.js';

test('stock image normalization preserves IDs and resolves legacy image fields', () => {
  const sourceLabels = [{ id: 'label-1', text: '保持', x: 25, y: 75, colorIndex: 2 }];
  const direct = normalizeStockImageEntry({ id: 'image-1', name: 'one.png', data: 'data:one', code: 'E1931', freeLabels: sourceLabels });
  const legacy = normalizeStockImageEntry({ imageId: 'image-2', originalName: 'two.png', image: 'data:two' });
  const referenced = normalizeStockImageEntry(
    { imageId: 'image-3', code: 'CODE-3' },
    { 'image-3': 'data:three' }
  );

  assert.equal(direct.id, 'image-1');
  assert.equal(direct.data, 'data:one');
  assert.equal(direct.code, 'E1931');
  assert.deepEqual(direct.freeLabels, sourceLabels);
  assert.notEqual(direct.freeLabels, sourceLabels);
  assert.equal(legacy.id, 'image-2');
  assert.equal(legacy.name, 'two.png');
  assert.equal(referenced.data, 'data:three');
  assert.equal(normalizeStockImageEntry({ id: 'missing' }), null);
});

test('stock image normalization removes duplicate IDs and duplicate image data', () => {
  const normalized = normalizeStockImages([
    { id: 'image-1', data: 'data:one' },
    { id: 'image-1', data: 'data:different' },
    { id: 'image-2', data: 'data:one' },
    { id: 'image-3', data: 'data:three' }
  ]);

  assert.deepEqual(normalized.map((image) => image.id), ['image-1', 'image-3']);
});

test('cloud image documents are converted into the current library format', () => {
  const documents = [
    {
      id: 'cloud-1',
      data: () => ({ name: 'E1682.jpg', code: 'E1682', data: 'data:image/jpeg;base64,one' })
    },
    {
      id: 'cloud-2',
      data: () => ({ name: 'E1690.jpg', image: 'data:image/jpeg;base64,two' })
    }
  ];

  const normalized = normalizeCloudImageDocuments(documents);

  assert.deepEqual(normalized.map((image) => image.id), ['cloud-1', 'cloud-2']);
  assert.deepEqual(normalized.map((image) => image.name), ['E1682.jpg', 'E1690.jpg']);
  assert.deepEqual(normalized.map((image) => image.data), [
    'data:image/jpeg;base64,one',
    'data:image/jpeg;base64,two'
  ]);
});

test('stock image list comparison ignores timestamps but detects identity changes', () => {
  const left = [{ id: 'image-1', name: 'one.png', data: 'data:one', freeLabels: [{ id: 'label-1', text: '保持' }], createdAt: { seconds: 1 } }];
  const same = [{ id: 'image-1', name: 'one.png', data: 'data:one', createdAt: { seconds: 2 } }];
  const changed = [{ id: 'image-1', name: 'one.png', data: 'data:changed', createdAt: { seconds: 1 } }];
  const changedLabel = [{ ...left[0], freeLabels: [{ id: 'label-1', text: '変更' }] }];

  assert.equal(isSameStockImageList(left, same), false);
  assert.equal(isSameStockImageList(left, [{ ...left[0], createdAt: { seconds: 2 } }]), true);
  assert.equal(isSameStockImageList(left, changed), false);
  assert.equal(isSameStockImageList(left, changedLabel), false);
});

test('workedBy is preserved by normalization and compared as identity', () => {
  const entry = normalizeStockImageEntry({ id: 'image-1', data: 'data:one', workedBy: ['user-a'] });
  assert.deepEqual(entry.workedBy, ['user-a']);
  // 未記録・空配列は null に正規化される
  assert.equal(normalizeStockImageEntry({ id: 'image-2', data: 'data:two' }).workedBy, null);
  assert.equal(normalizeStockImageEntry({ id: 'image-3', data: 'data:three', workedBy: [] }).workedBy, null);

  const base = [{ id: 'image-1', data: 'data:one', workedBy: ['user-a'] }];
  assert.equal(isSameStockImageList(base, [{ ...base[0] }]), true);
  assert.equal(isSameStockImageList(base, [{ ...base[0], workedBy: ['user-a', 'user-b'] }]), false);
});

test('PDF crop provenance is preserved and compared without affecting legacy images', () => {
  const source = {
    id: 'image-pdf',
    name: 'E1957.jpg',
    data: 'data:pdf',
    code: 'E1957',
    sourcePdfName: 'P010.pdf',
    sourcePage: 10,
    pdfPageNumber: 1,
    sizeType: '1/8 横（2コマ）',
    cropRect: { x: 0.1, y: 0.06, width: 0.4, height: 0.2 },
    sourceText: 'アイソカルゼリー 261-E1957',
    sourceTextVersion: 1,
    sourceTextTruncated: false,
    textExtractionRect: { x: 0.11, y: 0.07, width: 0.38, height: 0.18 },
    catalogCode: 'E1957',
    productName: 'アイソカルゼリー',
    productNameSource: 'csv',
    priceIncludingTax: 2139,
    priceExcludingTax: 1980,
    priceCandidates: [{ amount: 2139, taxType: 'including', text: '¥2,139' }],
    priceExtractionConfidence: 'high',
    catalogTextData: {
      version: 1,
      itemNumber: 'ABC-123',
      itemNumberCandidates: ['ABC-123'],
      catchCopy: 'おいしく栄養補給。',
      catchCopyCandidates: ['おいしく栄養補給。'],
      availability: 'stock',
      availabilityLabels: ['在庫商品'],
      handlingMarkers: ['(D)'],
      hasDemoMarker: true,
      specifications: ['●内容量/100g'],
      compositionDetails: [],
      materialDetails: []
    }
  };
  const normalized = normalizeStockImageEntry(source);
  assert.equal(normalized.sourcePdfName, 'P010.pdf');
  assert.equal(normalized.sourcePage, 10);
  assert.equal(normalized.sourceText, source.sourceText);
  assert.deepEqual(normalized.cropRect, source.cropRect);
  assert.notEqual(normalized.cropRect, source.cropRect);
  assert.deepEqual(normalized.textExtractionRect, source.textExtractionRect);
  assert.notEqual(normalized.textExtractionRect, source.textExtractionRect);
  assert.deepEqual(normalized.priceCandidates, source.priceCandidates);
  assert.notEqual(normalized.priceCandidates, source.priceCandidates);
  assert.equal(normalized.productName, 'アイソカルゼリー');
  assert.equal(normalized.productNameSource, 'csv');
  assert.equal(normalized.priceIncludingTax, 2139);
  assert.deepEqual(normalized.catalogTextData, source.catalogTextData);
  assert.notEqual(normalized.catalogTextData, source.catalogTextData);
  assert.notEqual(normalized.catalogTextData.specifications, source.catalogTextData.specifications);
  assert.equal(isSameStockImageList([source], [{ ...source }]), true);
  assert.equal(isSameStockImageList([source], [{ ...source, sourcePage: 11 }]), false);
});

test('worked-by filter shows own and legacy images only', () => {
  // 未記録 (既存データ) は互換のため全員に表示
  assert.equal(isImageWorkedByUser({ id: 'legacy' }, 'user-a'), true);
  assert.equal(isImageWorkedByUser({ id: 'legacy', workedBy: [] }, 'user-a'), true);
  // 記録済みは作業者のみ
  assert.equal(isImageWorkedByUser({ workedBy: ['user-a'] }, 'user-a'), true);
  assert.equal(isImageWorkedByUser({ workedBy: ['user-b'] }, 'user-a'), false);
  assert.equal(isImageWorkedByUser({ workedBy: ['user-b', 'user-a'] }, 'user-a'), true);
  // ユーザー未確定時は隠さない
  assert.equal(isImageWorkedByUser({ workedBy: ['user-b'] }, null), true);
});

// --- 画像削除の同一性判定 ---
import {
  buildImageDeletionIdentity,
  collectSheetImageKeys,
  createImageDeletionFilter,
  getImageDeletionMatchKind,
  getStockImageCode
} from '../src/domain/images.js';

test('getStockImageCode reads the product code from the code field or filename stem', () => {
  assert.equal(getStockImageCode({ code: 'B0867' }), 'B0867');
  assert.equal(getStockImageCode({ name: 'B0867.jpg' }), 'B0867');
  assert.equal(getStockImageCode({ name: '261-B0867.jpg' }), 'B0867');
  assert.equal(getStockImageCode({ originalName: 'ｅ１９３１.png' }), 'E1931');
  // 商品コードに見えない名前は空 (一般名のファイルを同一視しない)
  assert.equal(getStockImageCode({ name: 'logo.png' }), '');
  assert.equal(getStockImageCode({ name: '集合写真.jpg' }), '');
  assert.equal(getStockImageCode({}), '');
});

test('buildImageDeletionIdentity resolves code and data from the library entry by id', () => {
  const images = [
    { id: 'img-1', name: 'B0867.jpg', code: 'B0867', data: 'data:a' },
    { id: 'img-2', name: 'E1931.jpg', data: 'data:b' }
  ];
  const identity = buildImageDeletionIdentity(['img-1', { id: 'img-2' }], images);
  assert.deepEqual([...identity.ids], ['img-1', 'img-2']);
  assert.deepEqual([...identity.data], ['data:a', 'data:b']);
  assert.deepEqual([...identity.codes], ['B0867', 'E1931']);
});

test('getImageDeletionMatchKind distinguishes direct hits from code-only duplicates', () => {
  const identity = buildImageDeletionIdentity([{ id: 'img-1', data: 'data:a', name: 'B0867.jpg' }], []);
  assert.equal(getImageDeletionMatchKind({ id: 'img-1' }, identity), 'direct');
  // 隠れた同一データの複製 (id 違い) も直接扱い
  assert.equal(getImageDeletionMatchKind({ id: 'twin', data: 'data:a' }, identity), 'direct');
  // 同じコードで別バイトの複製はコード一致
  assert.equal(getImageDeletionMatchKind({ id: 'other', name: 'B0867.jpg', data: 'data:z' }, identity), 'code');
  assert.equal(getImageDeletionMatchKind({ id: 'x', name: 'E9999.jpg', data: 'data:y' }, identity), null);
});

test('createImageDeletionFilter keeps code-only duplicates that are placed on a sheet', () => {
  const identity = buildImageDeletionIdentity([{ id: 'img-1', data: 'data:a', name: 'B0867.jpg' }], []);
  const sheetKeys = collectSheetImageKeys([
    { panels: [{ imageId: 'placed-twin' }, { image: 'data:placed' }, null] }
  ]);
  assert.deepEqual([...sheetKeys.ids], ['placed-twin']);
  assert.deepEqual([...sheetKeys.data], ['data:placed']);

  const kindOf = createImageDeletionFilter({ identity, sheetKeys });
  // 未配置のコード一致複製は削除対象
  assert.equal(kindOf({ id: 'free-twin', name: 'B0867.jpg', data: 'data:z' }), 'code');
  // 配置中のコード一致複製は守る (コマの表示が壊れるため)
  assert.equal(kindOf({ id: 'placed-twin', name: 'B0867.jpg', data: 'data:w' }), null);
  assert.equal(kindOf({ id: 'p2', name: 'B0867.jpg', data: 'data:placed' }), null);
  // 直接指定は配置中でも削除 (ユーザーが明示的に選んだもの)
  assert.equal(kindOf({ id: 'img-1' }), 'direct');
  assert.equal(kindOf({ id: 'unrelated', name: 'E1.png', data: 'data:u' }), null);
});
