import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';

import {
  clearActiveNativeDragPayload,
  extractPanelAssignmentFromDragPayload,
  extractPanelMoveDragPayload,
  getActiveNativeDragPayload,
  getActivePanelMoveDragPayload,
  getDragPayload,
  isDropEventHandled,
  markDropEventHandled,
  normalizeDragPayload,
  setDragPayload
} from '../src/lib/dragPayload.js';

const createDataTransfer = () => {
  const values = new Map();
  return {
    getData(type) {
      return values.get(type) || '';
    },
    setData(type, value) {
      values.set(type, value);
    }
  };
};

afterEach(() => {
  clearActiveNativeDragPayload();
});

test('drag payload normalization keeps a stable string-based shape', () => {
  const normalized = normalizeDragPayload({ sourceIndex: 2, imageId: null, isText: false });

  assert.equal(normalized.__daiwariDragPayload, true);
  assert.equal(normalized.sourceIndex, '2');
  assert.equal(normalized.imageId, '');
  assert.equal(normalized.isText, 'false');
  assert.equal(normalized.label, '');
  assert.equal(normalized.freeLabels, '');
});

test('native drag payload round-trips and preserves the active panel move', () => {
  const dataTransfer = createDataTransfer();
  setDragPayload(dataTransfer, {
    moveSourceType: 'panel',
    sourceSheetId: 'sheet-1',
    sourceIndex: 3,
    textData: 'draft'
  });

  const payload = getDragPayload(dataTransfer);
  assert.equal(payload.sourceSheetId, 'sheet-1');
  assert.equal(getActiveNativeDragPayload().sourceIndex, '3');
  assert.deepEqual(getActivePanelMoveDragPayload(), {
    sourceSheetId: 'sheet-1',
    sourceIndex: 3,
    movedText: 'draft'
  });
});

test('panel move extraction rejects incomplete and non-panel payloads', () => {
  assert.equal(extractPanelMoveDragPayload({ moveSourceType: 'stock' }), null);
  assert.equal(extractPanelMoveDragPayload({ moveSourceType: 'panel', sourceIndex: '2' }), null);
  assert.deepEqual(
    extractPanelMoveDragPayload({
      moveSourceType: 'panel',
      sourceSheetId: 'sheet-2',
      sourceIndex: '4',
      textData: 'text'
    }),
    { sourceSheetId: 'sheet-2', sourceIndex: 4, movedText: 'text' }
  );
});

test('panel assignment retains text payloads and filename code fallback', () => {
  assert.deepEqual(
    extractPanelAssignmentFromDragPayload({
      src: 'data:image/png;base64,image',
      name: 'a1234-product.png'
    }),
    {
      image: 'data:image/png;base64,image',
      imageId: null,
      label: null,
      code: 'A1234',
      isText: false,
      text: '',
      freeLabels: [],
      freeText: null,
      fromTempId: null,
      fromExcludedId: null
    }
  );
  assert.equal(
    extractPanelAssignmentFromDragPayload({
      isText: 'true',
      hasTextPayload: '1',
      textPayload: 'transferred'
    }, 'fallback').text,
    'transferred'
  );
  assert.equal(extractPanelAssignmentFromDragPayload({}, 'fallback'), null);
});

test('temporary item drag payload round-trips free labels and legacy free text', () => {
  const dataTransfer = createDataTransfer();
  setDragPayload(dataTransfer, {
    src: 'data:image/png;base64,image',
    imageId: 'image-1',
    fromTempId: 'temp-1',
    freeLabels: [{ id: 'note-1', text: '持ち運ぶ', x: 30, y: 40, colorIndex: 3 }]
  });

  const assignment = extractPanelAssignmentFromDragPayload(getDragPayload(dataTransfer));
  assert.deepEqual(assignment.freeLabels, [
    { id: 'note-1', text: '持ち運ぶ', x: 30, y: 40, colorIndex: 3 }
  ]);
  assert.equal(assignment.freeText, null);

  const legacyAssignment = extractPanelAssignmentFromDragPayload({
    src: 'data:image/png;base64,legacy',
    freeLabels: 'not-json',
    freeText: '旧ラベル'
  });
  assert.deepEqual(legacyAssignment.freeLabels, [
    { id: 'legacy', text: '旧ラベル', x: 50, y: 50, colorIndex: 0 }
  ]);
});

test('drop events can be marked as handled exactly once', () => {
  const nativeEvent = {};
  assert.equal(isDropEventHandled(nativeEvent), false);
  markDropEventHandled(nativeEvent);
  assert.equal(isDropEventHandled(nativeEvent), true);
});
