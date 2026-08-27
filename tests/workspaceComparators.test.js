import test from 'node:test';
import assert from 'node:assert/strict';

import {
  isSameSheetList,
  isSameTransferItemList,
  toComparableSeconds
} from '../src/domain/workspaceComparators.js';

test('Firestore-like timestamps are converted to comparable seconds', () => {
  assert.equal(toComparableSeconds({ seconds: 123 }), 123);
  assert.equal(toComparableSeconds({ toDate: () => new Date(456000) }), 456);
  assert.equal(toComparableSeconds(789000), 789);
  assert.equal(toComparableSeconds(null), 0);
  assert.equal(toComparableSeconds({ toDate: () => { throw new Error('invalid'); } }), 0);
});

test('transfer item comparison covers persisted identity and content fields', () => {
  const left = [{
    id: 'item-1',
    imageId: 'image-1',
    label: 'label',
    code: 'code',
    text: 'text',
    isText: true,
    originalName: 'one.png',
    createdAt: { seconds: 10 }
  }];
  const same = [{ ...left[0], createdAt: { toDate: () => new Date(10000) } }];
  const changed = [{ ...left[0], text: 'changed' }];

  assert.equal(isSameTransferItemList(left, same), true);
  assert.equal(isSameTransferItemList(left, changed), false);
});

test('sheet comparison uses the existing panel comparison contract', () => {
  const left = [{
    id: 'sheet-1',
    genre: 'meal',
    order: 1,
    panels: [{ code: 'code-1', fromTempId: 'before' }]
  }];
  const same = [{
    id: 'sheet-1',
    genre: 'meal',
    order: 1,
    panels: [{ code: 'code-1', fromTempId: 'after' }]
  }];
  const changed = [{ ...same[0], genre: 'bath' }];

  assert.equal(isSameSheetList(left, same), true);
  assert.equal(isSameSheetList(left, changed), false);
});
