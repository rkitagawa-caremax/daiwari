import test from 'node:test';
import assert from 'node:assert/strict';

import {
  DEFAULT_PANEL_DATA,
  PANEL_COUNT,
  applyPanelTransferableContent,
  buildDefaultPanels,
  buildPanelMapUpdates,
  clearPanelTransferableContent,
  getPanelDataPatch,
  getPanelCsvCode,
  getPanelFreeLabels,
  getPanelsFromDocData,
  getPanelTransferableContent,
  hasPanelTransferableContent,
  isPanelDataEqual,
  sanitizePanelData,
  swapPanelTransferableContent,
  toPanelsMap
} from '../src/domain/panels.js';

test('panel comparison preserves the existing transient-field and null rules', () => {
  assert.equal(
    isPanelDataEqual(
      { text: undefined, fromTempId: 'before', fromExcludedId: 'before' },
      { text: null, fromTempId: 'after', fromExcludedId: 'after' }
    ),
    true
  );
  assert.equal(isPanelDataEqual({ meta: { x: 1 } }, { meta: { x: 1 } }), true);
  assert.equal(isPanelDataEqual({ meta: { x: 1 } }, { meta: { x: 2 } }), false);
});

test('panel patch contains changed values and nulls for removed fields', () => {
  const patch = getPanelDataPatch(
    { image: 'old', text: 'same', meta: { x: 1 } },
    { text: 'same', meta: { x: 2 } }
  );

  assert.deepEqual(patch, {
    image: null,
    meta: { x: 2 }
  });
});

test('panel transfer content carries free labels while preserving target layout', () => {
  const source = {
    image: 'data:image/png;base64,source',
    imageId: 'image-1',
    code: 'A1000',
    text: 'source text',
    rowSpan: 4,
    colSpan: 1,
    freeLabels: [{ id: 'label-1', text: '注記', x: 25, y: 75, colorIndex: 2 }]
  };
  const target = {
    rowSpan: 2,
    colSpan: 3,
    freeLabels: [{ id: 'old', text: '古いラベル', x: 50, y: 50, colorIndex: 0 }]
  };

  const transferred = getPanelTransferableContent(source, 'edited text');
  const assigned = applyPanelTransferableContent(target, source, 'edited text');
  const cleared = clearPanelTransferableContent(source);

  assert.deepEqual(transferred.freeLabels, source.freeLabels);
  assert.notEqual(transferred.freeLabels, source.freeLabels);
  assert.notEqual(transferred.freeLabels[0], source.freeLabels[0]);
  assert.equal(transferred.text, 'edited text');
  assert.equal(transferred.code, 'A1000');
  assert.equal(assigned.rowSpan, 2);
  assert.equal(assigned.colSpan, 3);
  assert.deepEqual(assigned.freeLabels, source.freeLabels);
  assert.deepEqual(cleared.freeLabels, []);
  assert.equal(cleared.freeText, null);
  assert.equal(cleared.rowSpan, 4);
  assert.equal(cleared.colSpan, 1);
  assert.equal(hasPanelTransferableContent({ freeLabels: source.freeLabels }), true);

  const assignmentWithoutLabels = applyPanelTransferableContent(target, { image: 'replacement' });
  assert.deepEqual(assignmentWithoutLabels.freeLabels, []);
});

test('arrange-mode swap preserves both panel layouts and both transferable payloads', () => {
  const source = {
    imageId: 'image-source',
    code: 'E1001',
    rowSpan: 2,
    colSpan: 1,
    freeLabels: [{ id: 'source-label', text: '移動元', x: 40, y: 60, colorIndex: 1 }]
  };
  const target = {
    imageId: 'image-target',
    code: 'E1002',
    rowSpan: 1,
    colSpan: 2,
    freeLabels: [{ id: 'target-label', text: '移動先', x: 55, y: 45, colorIndex: 2 }]
  };

  const swapped = swapPanelTransferableContent(source, target);

  assert.equal(swapped.sourcePanel.rowSpan, 2);
  assert.equal(swapped.sourcePanel.colSpan, 1);
  assert.equal(swapped.sourcePanel.imageId, 'image-target');
  assert.deepEqual(swapped.sourcePanel.freeLabels, target.freeLabels);
  assert.equal(swapped.targetPanel.rowSpan, 1);
  assert.equal(swapped.targetPanel.colSpan, 2);
  assert.equal(swapped.targetPanel.imageId, 'image-source');
  assert.deepEqual(swapped.targetPanel.freeLabels, source.freeLabels);

  const movedToEmpty = swapPanelTransferableContent(source, {
    rowSpan: 1,
    colSpan: 1,
    text: 'stale non-content text'
  });
  assert.equal(movedToEmpty.sourcePanel.imageId, null);
  assert.equal(movedToEmpty.sourcePanel.text, '');
  assert.equal(movedToEmpty.targetPanel.imageId, 'image-source');
});

test('CSV code keeps text dummy codes while retaining legacy dummy markers', () => {
  assert.equal(getPanelCsvCode({ label: 'テキスト', isText: true, code: 'E1931' }), 'E1931');
  assert.equal(getPanelCsvCode({ label: 'テキスト', isText: true, code: null }), '');
  assert.equal(getPanelCsvCode({ label: 'タイトル', code: 'E1931' }), 'ダミーコマ');
  assert.equal(getPanelCsvCode({ code: 'E1931' }), 'E1931');
});

test('legacy free text is normalized into a movable free label', () => {
  assert.deepEqual(getPanelFreeLabels({ freeText: '旧ラベル' }), [
    { id: 'legacy', text: '旧ラベル', x: 50, y: 50, colorIndex: 0 }
  ]);
  assert.deepEqual(getPanelTransferableContent({ freeText: '旧ラベル' }).freeLabels, [
    { id: 'legacy', text: '旧ラベル', x: 50, y: 50, colorIndex: 0 }
  ]);
});

test('panel sanitization is shallow, nulls undefined, and prefers imageId', () => {
  const input = {
    image: 'data:image/png;base64,old',
    imageId: 'image-1',
    optional: undefined,
    nested: { optional: undefined }
  };

  const sanitized = sanitizePanelData(input);

  assert.deepEqual(sanitized, {
    image: null,
    imageId: 'image-1',
    optional: null,
    nested: { optional: undefined }
  });
  assert.equal(input.image, 'data:image/png;base64,old');
});

test('default panels contain 16 independent panel objects', () => {
  const panels = buildDefaultPanels();

  assert.equal(panels.length, PANEL_COUNT);
  assert.deepEqual(panels[0], DEFAULT_PANEL_DATA);
  assert.notEqual(panels[0], panels[1]);

  panels[0].text = 'changed';
  assert.equal(panels[1].text, '');
});

test('panelsMap overrides legacy panels while invalid map entries are ignored', () => {
  const panels = getPanelsFromDocData({
    panels: [
      { text: 'legacy text' },
      { code: 'legacy-code' }
    ],
    panelsMap: {
      0: { text: 'map text' },
      2: [],
      3: { rowSpan: 2 },
      16: { text: 'out of range' },
      invalid: { text: 'invalid index' }
    }
  });

  assert.equal(panels.length, PANEL_COUNT);
  assert.equal(panels[0].text, 'map text');
  assert.equal(panels[1].code, 'legacy-code');
  assert.deepEqual(panels[2], DEFAULT_PANEL_DATA);
  assert.equal(panels[3].rowSpan, 2);
});

test('toPanelsMap creates exactly 16 canonical and sanitized entries', () => {
  const panelsMap = toPanelsMap([
    {
      image: 'data:image/png;base64,old',
      imageId: 'image-1',
      custom: undefined
    }
  ]);

  assert.deepEqual(Object.keys(panelsMap), Array.from({ length: PANEL_COUNT }, (_, index) => String(index)));
  assert.equal(panelsMap['0'].image, null);
  assert.equal(panelsMap['0'].imageId, 'image-1');
  assert.equal(panelsMap['0'].custom, null);
  assert.deepEqual(panelsMap['15'], DEFAULT_PANEL_DATA);
});

test('panel map updates contain only changed full panel entries', () => {
  const previousPanels = buildDefaultPanels();
  const nextPanels = buildDefaultPanels();
  nextPanels[2] = {
    ...nextPanels[2],
    code: 'ABC-123'
  };

  const updates = buildPanelMapUpdates(previousPanels, nextPanels);

  assert.deepEqual(Object.keys(updates), ['panelsMap.2']);
  assert.deepEqual(updates['panelsMap.2'], {
    ...DEFAULT_PANEL_DATA,
    code: 'ABC-123'
  });

  const transientOnly = buildPanelMapUpdates(
    previousPanels,
    previousPanels.map((panel, index) => (
      index === 2 ? { ...panel, fromTempId: 'temporary-id' } : panel
    ))
  );
  assert.deepEqual(transientOnly, {});
});

test('canonical panelsMap round-trips through the document reader', () => {
  const input = buildDefaultPanels();
  input[4] = {
    ...input[4],
    image: 'data:image/png;base64,old',
    imageId: 'image-4',
    rowSpan: 2
  };

  const restored = getPanelsFromDocData({ panelsMap: toPanelsMap(input) });

  assert.equal(restored.length, PANEL_COUNT);
  assert.equal(restored[4].image, null);
  assert.equal(restored[4].imageId, 'image-4');
  assert.equal(restored[4].rowSpan, 2);
});
