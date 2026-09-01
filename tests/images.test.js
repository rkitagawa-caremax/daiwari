import test from 'node:test';
import assert from 'node:assert/strict';

import {
  isImageWorkedByUser,
  isSameStockImageList,
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
    sourceTextTruncated: false
  };
  const normalized = normalizeStockImageEntry(source);
  assert.equal(normalized.sourcePdfName, 'P010.pdf');
  assert.equal(normalized.sourcePage, 10);
  assert.equal(normalized.sourceText, source.sourceText);
  assert.deepEqual(normalized.cropRect, source.cropRect);
  assert.notEqual(normalized.cropRect, source.cropRect);
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
