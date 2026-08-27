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
