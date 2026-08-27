import test from 'node:test';
import assert from 'node:assert/strict';

import {
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
