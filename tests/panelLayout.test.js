import test from 'node:test';
import assert from 'node:assert/strict';

import {
  canPlacePanelAt,
  fillPanelArea,
  findFirstPlaceableIndex,
  getCoords,
  getSizeType,
  getSpansFromSizeTypeRobust
} from '../src/domain/panelLayout.js';

test('panel sizes retain their display labels and robust reverse mapping', () => {
  assert.equal(getSizeType(1, 1), '1/16（1コマ）');
  assert.equal(getSizeType(2, 3), '6/16 横（6コマ）');
  assert.equal(getSizeType(4, 4), '1P（16コマ）');
  assert.equal(getSizeType(3, 3), 'custom');

  assert.deepEqual(getSpansFromSizeTypeRobust('１／８ 縦（２コマ）'), { r: 2, c: 1 });
  assert.deepEqual(getSpansFromSizeTypeRobust('1/4 横（4コマ）'), { r: 1, c: 4 });
  assert.deepEqual(getSpansFromSizeTypeRobust('8コマ vertical'), { r: 4, c: 2 });
  assert.deepEqual(getSpansFromSizeTypeRobust('unknown'), { r: 1, c: 1 });
});

test('grid coordinates and placement respect the 4 by 4 boundary', () => {
  assert.deepEqual(getCoords(0), { row: 0, col: 0 });
  assert.deepEqual(getCoords(15), { row: 3, col: 3 });
  assert.equal(canPlacePanelAt(0, 2, 2, new Set()), true);
  assert.equal(canPlacePanelAt(3, 1, 2, new Set()), false);
  assert.equal(canPlacePanelAt(0, 2, 2, new Set([5])), false);
});

test('first-place search wraps to the beginning and reports no available slot', () => {
  assert.equal(findFirstPlaceableIndex(1, 1, new Set([5, 6]), 5), 7);
  assert.equal(findFirstPlaceableIndex(1, 1, new Set([5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]), 5), 0);
  assert.equal(findFirstPlaceableIndex(1, 1, new Set(Array.from({ length: 16 }, (_, index) => index))), -1);
});

test('filling a panel area marks covered slots hidden and occupied', () => {
  const panels = Array.from({ length: 16 }, () => ({ hidden: false }));
  const occupied = new Set();

  fillPanelArea(panels, 1, 2, 2, occupied);

  assert.deepEqual(Array.from(occupied).sort((left, right) => left - right), [1, 2, 5, 6]);
  assert.equal(panels[1].hidden, false);
  assert.equal(panels[2].hidden, true);
  assert.equal(panels[5].hidden, true);
  assert.equal(panels[6].hidden, true);
});
