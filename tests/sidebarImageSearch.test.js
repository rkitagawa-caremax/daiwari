import test from 'node:test';
import assert from 'node:assert/strict';

import { buildSidebarImageResults } from '../src/domain/sidebarImageSearch.js';

const images = [
  { id: 'assigned-image', name: 'E1931.png', data: 'data:assigned' },
  { id: 'available-image', name: 'E2000.png', data: 'data:available', code: 'E2000' },
  { id: 'excluded-image', name: 'E3000.png', data: 'data:excluded' }
];

const sheets = [
  {
    id: 'sheet-meal',
    genre: 'meal',
    panels: [
      { imageId: 'assigned-image', code: 'E1931', freeLabels: [{ id: 'label-1', text: '保持' }] }
    ]
  }
];

const excludedItems = [{ imageId: 'excluded-image' }];

test('sidebar default image list keeps assigned and excluded images hidden', () => {
  const result = buildSidebarImageResults({ images, sheets, excludedItems });

  assert.deepEqual(result.map((image) => image.id), ['available-image']);
  assert.equal(result[0].assignment, undefined);
});

test('sidebar code search includes assigned image with page and genre metadata', () => {
  const result = buildSidebarImageResults({
    images,
    sheets,
    excludedItems,
    searchQuery: 'e1931'
  });

  assert.equal(result.length, 1);
  assert.equal(result[0].id, 'assigned-image');
  assert.equal(result[0].data, 'data:assigned');
  assert.equal(result[0].code, 'E1931');
  assert.deepEqual(result[0].freeLabels, [{ id: 'label-1', text: '保持' }]);
  assert.deepEqual(result[0].assignment, {
    sheetId: 'sheet-meal',
    sheetNumber: 1,
    genre: 'meal',
    panelIndex: 0,
    code: 'E1931'
  });
});

test('sidebar search still finds available images by code and filename', () => {
  const byCode = buildSidebarImageResults({ images, sheets, excludedItems, searchQuery: 'E2000' });
  const byFilename = buildSidebarImageResults({ images, sheets, excludedItems, searchQuery: '2000.png' });

  assert.deepEqual(byCode.map((image) => image.id), ['available-image']);
  assert.deepEqual(byFilename.map((image) => image.id), ['available-image']);
});

test('sidebar assigned search resolves legacy panel image data without a stock record', () => {
  const result = buildSidebarImageResults({
    images: [],
    sheets: [{
      id: 'legacy-sheet',
      genre: 'bath',
      panels: [{ image: 'data:legacy', code: 'E9999', originalName: 'legacy.png' }]
    }],
    searchQuery: 'E9999'
  });

  assert.equal(result.length, 1);
  assert.equal(result[0].data, 'data:legacy');
  assert.equal(result[0].assignment.sheetId, 'legacy-sheet');
  assert.equal(result[0].assignment.sheetNumber, 1);
});

test('sidebar image filter narrows unassigned list but keeps assigned search results', () => {
  const workImages = [
    { id: 'mine', name: 'mine.png', data: 'data:mine', workedBy: ['user-a'] },
    { id: 'others', name: 'others.png', data: 'data:others', workedBy: ['user-b'] },
    { id: 'legacy', name: 'legacy.png', data: 'data:legacy-2' }
  ];
  const onlyMine = (image) => !image.workedBy || image.workedBy.includes('user-a');

  const filtered = buildSidebarImageResults({ images: workImages, imageFilter: onlyMine });
  assert.deepEqual(filtered.map((image) => image.id), ['mine', 'legacy']);

  // フィルタなし (ALL) は従来どおり全件
  const all = buildSidebarImageResults({ images: workImages });
  assert.deepEqual(all.map((image) => image.id), ['mine', 'others', 'legacy']);

  // 検索時、配置済みナビゲーション結果はフィルタの影響を受けない
  const searched = buildSidebarImageResults({
    images: [...images, ...workImages],
    sheets,
    excludedItems,
    searchQuery: 'e1931',
    imageFilter: onlyMine
  });
  assert.equal(searched.length, 1);
  assert.equal(searched[0].id, 'assigned-image');
});
