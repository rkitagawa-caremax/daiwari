import assert from 'node:assert/strict';
import test from 'node:test';

import {
  dispatchWorkspaceDrop,
  parsePanelDropZone
} from '../src/lib/workspaceDropZones.js';

const createDropHandlers = () => {
  const calls = [];
  return {
    calls,
    handlers: {
      onDropToPanel: (...args) => {
        calls.push(['panel', ...args]);
        return true;
      },
      onDropToTemp: (...args) => {
        calls.push(['temp', ...args]);
        return true;
      },
      onDropToStock: (...args) => {
        calls.push(['stock', ...args]);
        return true;
      },
      onDropToExcluded: (...args) => {
        calls.push(['excluded', ...args]);
        return true;
      }
    }
  };
};

test('parsePanelDropZone extracts sheet id and panel index', () => {
  assert.deepEqual(parsePanelDropZone('panel:sheet-a:7'), {
    sheetId: 'sheet-a',
    panelIndex: 7
  });
});

test('parsePanelDropZone supports sheet ids containing colons', () => {
  assert.deepEqual(parsePanelDropZone('panel:group:sheet-a:12'), {
    sheetId: 'group:sheet-a',
    panelIndex: 12
  });
});

test('parsePanelDropZone rejects non-panel and malformed zones', () => {
  assert.equal(parsePanelDropZone('temp'), null);
  assert.equal(parsePanelDropZone('panel:sheet-a'), null);
  assert.equal(parsePanelDropZone('panel::2'), null);
  assert.equal(parsePanelDropZone('panel:sheet-a:not-a-number'), null);
});

test('dispatchWorkspaceDrop routes both page panel zones to the panel handler', () => {
  const { calls, handlers } = createDropHandlers();
  const dragPayload = { imageId: 'image-1' };

  assert.equal(dispatchWorkspaceDrop({
    zoneId: 'panel:left-page:3',
    dragPayload,
    ...handlers
  }), true);
  assert.equal(dispatchWorkspaceDrop({
    zoneId: 'panel:right-page:9',
    dragPayload,
    ...handlers
  }), true);

  assert.deepEqual(calls, [
    ['panel', 'left-page', 3, dragPayload],
    ['panel', 'right-page', 9, dragPayload]
  ]);
});

test('dispatchWorkspaceDrop routes regular shelf zones', () => {
  const { calls, handlers } = createDropHandlers();
  const dragPayload = { imageId: 'image-1' };

  ['temp', 'stock', 'excluded'].forEach((zoneId) => {
    assert.equal(dispatchWorkspaceDrop({ zoneId, dragPayload, ...handlers }), true);
  });

  assert.deepEqual(calls.map(([target]) => target), ['temp', 'stock', 'excluded']);
});

test('panel arrange payloads cannot leave the page workspace', () => {
  const { calls, handlers } = createDropHandlers();
  const dragPayload = {
    arrangeMode: 'true',
    arrangeSheetId: 'left-page',
    arrangeTokenId: 'token-1'
  };

  assert.equal(dispatchWorkspaceDrop({
    zoneId: 'stock',
    dragPayload,
    ...handlers
  }), false);
  assert.equal(calls.length, 0);

  assert.equal(dispatchWorkspaceDrop({
    zoneId: 'panel:right-page:4',
    dragPayload,
    ...handlers
  }), true);
  assert.deepEqual(calls, [['panel', 'right-page', 4, dragPayload]]);
});

test('dispatchWorkspaceDrop ignores unknown and malformed zones', () => {
  const { calls, handlers } = createDropHandlers();

  assert.equal(dispatchWorkspaceDrop({ zoneId: '', ...handlers }), false);
  assert.equal(dispatchWorkspaceDrop({ zoneId: 'unknown', ...handlers }), false);
  assert.equal(dispatchWorkspaceDrop({ zoneId: 'panel:sheet-a:nope', ...handlers }), false);
  assert.equal(calls.length, 0);
});
