import test from 'node:test';
import assert from 'node:assert/strict';

import {
  DEFAULT_PANEL_DATA,
  PANEL_COUNT,
  buildDefaultPanels,
  buildPanelMapUpdates,
  getPanelDataPatch,
  getPanelsFromDocData,
  isPanelDataEqual,
  sanitizePanelData,
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
