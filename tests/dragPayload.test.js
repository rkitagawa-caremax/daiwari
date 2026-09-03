import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';

import {
  PANEL_ARRANGE_HOLD_MS,
  PANEL_ARRANGE_MOVE_TOLERANCE_PX,
  buildExcludedItemDragConfig,
  buildLibraryImageDragConfig,
  buildPanelDragConfig,
  clearActiveNativeDragPayload,
  extractPanelAssignmentFromDragPayload,
  extractPanelArrangeDragPayload,
  extractPanelMoveDragPayload,
  getActiveNativeDragPayload,
  getActivePanelMoveDragPayload,
  getDragPayload,
  hasPanelArrangeHoldMoved,
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

test('panel drag config keeps native payload and pointer preview values aligned', () => {
  assert.deepEqual(buildPanelDragConfig({
    sheetId: 'sheet-1',
    panelIndex: 3,
    textData: '編集中テキスト',
    arrangeMode: true,
    arrangeSheetId: 'sheet-1',
    arrangeTokenId: 'token-1',
    previewImage: 'data:image/png;base64,panel',
    previewLabel: 'テキスト',
    previewCode: 'E1234',
    previewText: 'プレビュー本文'
  }), {
    payload: {
      moveSourceType: 'panel',
      sourceSheetId: 'sheet-1',
      sourceIndex: 3,
      textData: '編集中テキスト',
      arrangeMode: true,
      arrangeSheetId: 'sheet-1',
      arrangeTokenId: 'token-1'
    },
    preview: {
      image: 'data:image/png;base64,panel',
      label: 'テキスト',
      code: 'E1234',
      text: 'プレビュー本文'
    }
  });

  assert.deepEqual(buildPanelDragConfig({
    sheetId: 'sheet-2',
    panelIndex: 4,
    arrangeMode: false,
    arrangeSheetId: '',
    arrangeTokenId: ''
  }).payload, {
    moveSourceType: 'panel',
    sourceSheetId: 'sheet-2',
    sourceIndex: 4,
    textData: '',
    arrangeMode: false,
    arrangeSheetId: '',
    arrangeTokenId: ''
  });
});

test('library image drag config preserves free-label arrays for both drag paths', () => {
  const freeLabels = [{ id: 'label-1', text: '注記', x: 20, y: 30, colorIndex: 2 }];
  const config = buildLibraryImageDragConfig({
    id: 'image-1',
    data: 'data:image/png;base64,library',
    name: 'E2000.png',
    code: 'E2000',
    freeLabels,
    freeText: '旧注記'
  });

  assert.deepEqual(config, {
    payload: {
      src: 'data:image/png;base64,library',
      imageId: 'image-1',
      type: 'image',
      name: 'E2000.png',
      code: 'E2000',
      freeLabels,
      freeText: '旧注記'
    },
    preview: {
      image: 'data:image/png;base64,library',
      code: 'E2000'
    }
  });
  assert.equal(config.payload.freeLabels, freeLabels);
});

test('excluded item drag config preserves text flags, source id, and hidden non-text metadata', () => {
  const freeLabels = [{ id: 'label-2', text: '除外注記', x: 40, y: 50, colorIndex: 1 }];
  const textConfig = buildExcludedItemDragConfig({
    id: 'excluded-1',
    imageId: 'image-2',
    originalName: 'E3000.png',
    label: 'テキスト',
    code: 'E3000',
    isText: true,
    text: '移動する本文',
    freeLabels,
    freeText: '旧ラベル'
  }, 'data:image/png;base64,excluded');

  assert.deepEqual(textConfig, {
    payload: {
      src: 'data:image/png;base64,excluded',
      type: 'image',
      name: 'E3000.png',
      label: 'テキスト',
      code: 'E3000',
      isText: 'true',
      hasTextPayload: '1',
      textPayload: '移動する本文',
      text: '移動する本文',
      freeLabels,
      freeText: '旧ラベル',
      fromExcludedId: 'excluded-1',
      imageId: 'image-2'
    },
    preview: {
      image: 'data:image/png;base64,excluded',
      label: 'テキスト',
      code: 'E3000',
      text: '移動する本文'
    }
  });

  const imageConfig = buildExcludedItemDragConfig({
    id: 'excluded-2',
    isText: false,
    text: '保持するがプレビューしない本文'
  });
  assert.equal(imageConfig.payload.isText, 'false');
  assert.equal(imageConfig.payload.hasTextPayload, '1');
  assert.equal(imageConfig.payload.textPayload, '保持するがプレビューしない本文');
  assert.equal(imageConfig.payload.fromExcludedId, 'excluded-2');
  assert.deepEqual(imageConfig.payload.freeLabels, []);
  assert.equal(imageConfig.preview.text, '');
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

test('panel arrange payload requires its explicit mode and owning sheet', () => {
  assert.equal(PANEL_ARRANGE_HOLD_MS, 2000);
  assert.equal(PANEL_ARRANGE_MOVE_TOLERANCE_PX, 4);
  assert.equal(extractPanelArrangeDragPayload({ arrangeMode: true }), null);
  assert.equal(extractPanelArrangeDragPayload({ arrangeSheetId: 'sheet-1' }), null);
  assert.deepEqual(
    extractPanelArrangeDragPayload({
      arrangeMode: 'true',
      arrangeSheetId: 'sheet-1',
      arrangeTokenId: 'token-1'
    }),
    { sheetId: 'sheet-1', tokenId: 'token-1' }
  );
});

test('panel arrange hold is cancelled at the movement tolerance boundary', () => {
  assert.equal(hasPanelArrangeHoldMoved(10, 10, 13, 10), false);
  assert.equal(hasPanelArrangeHoldMoved(10, 10, 14, 10), true);
  assert.equal(hasPanelArrangeHoldMoved(10, 10, 12, 13), false);
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
