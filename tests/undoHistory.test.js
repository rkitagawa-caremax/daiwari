import test from 'node:test';
import assert from 'node:assert/strict';

import {
  applyUndoDomainChanges,
  buildUndoDomainChanges,
  hasUndoEntryChanges,
  mergeUndoDomainChanges
} from '../src/domain/undoHistory.js';

const isSameItem = (left, right) => left.id === right.id && left.value === right.value;

test('undo changes capture additions, edits, deletions, and positions', () => {
  const before = [
    { id: 'a', value: 1 },
    { id: 'b', value: 2 }
  ];
  const after = [
    { id: 'a', value: 3 },
    { id: 'c', value: 4 }
  ];

  const changes = buildUndoDomainChanges(before, after, isSameItem);
  assert.deepEqual(changes.map((change) => change.id).sort(), ['a', 'b', 'c']);
  assert.equal(changes.find((change) => change.id === 'a').before.value, 1);
  assert.equal(changes.find((change) => change.id === 'b').after, null);
  assert.equal(changes.find((change) => change.id === 'c').before, null);
  assert.deepEqual(applyUndoDomainChanges(after, changes), before);
});

test('adding an item does not mark every shifted item as changed', () => {
  const before = [
    { id: 'a', value: 1 },
    { id: 'b', value: 2 }
  ];
  const after = [
    { id: 'c', value: 3 },
    ...before
  ];

  const changes = buildUndoDomainChanges(before, after, isSameItem);
  assert.deepEqual(changes.map((change) => change.id), ['c']);
  assert.deepEqual(applyUndoDomainChanges(after, changes), before);
});

test('grouped undo keeps the earliest value and latest value', () => {
  const first = buildUndoDomainChanges(
    [{ id: 'a', value: 1 }],
    [{ id: 'a', value: 2 }],
    isSameItem
  );
  const second = buildUndoDomainChanges(
    [{ id: 'a', value: 2 }],
    [{ id: 'a', value: 3 }],
    isSameItem
  );
  const merged = mergeUndoDomainChanges(first, second, isSameItem);

  assert.equal(merged.length, 1);
  assert.equal(merged[0].before.value, 1);
  assert.equal(merged[0].after.value, 3);
});

test('grouped undo removes an edit that returned to its original value', () => {
  const first = buildUndoDomainChanges(
    [{ id: 'a', value: 1 }],
    [{ id: 'a', value: 2 }],
    isSameItem
  );
  const second = buildUndoDomainChanges(
    [{ id: 'a', value: 2 }],
    [{ id: 'a', value: 1 }],
    isSameItem
  );

  assert.deepEqual(mergeUndoDomainChanges(first, second, isSameItem), []);
});

test('undo entry reports whether any workspace domain changed', () => {
  assert.equal(hasUndoEntryChanges({ changes: { sheets: [] } }), false);
  assert.equal(hasUndoEntryChanges({ changes: { sheets: [{ id: 'a' }] } }), true);
});
