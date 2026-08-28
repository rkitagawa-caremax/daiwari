import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildPanelArrangeFinalPanels,
  buildPanelArrangeView,
  createPanelArrangeSession,
  getUnresolvedPanelArrangeTokens,
  isPanelArrangeSessionComplete,
  reconcilePanelArrangeSession,
  stagePanelArrangeDrop
} from '../src/domain/panelArrange.js';
import { buildDefaultPanels } from '../src/domain/panels.js';

const buildImagePanel = (base, imageId, code, labelText) => ({
  ...base,
  imageId,
  code,
  originalName: `${code}.png`,
  freeLabels: [{ id: `${code}-label`, text: labelText, x: 50, y: 50, colorIndex: 0 }]
});

test('arrange session captures image metadata and starts fully assigned', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel(panels[0], 'image-a', 'E1001', 'Aラベル');
  panels[1] = buildImagePanel(panels[1], 'image-b', 'E1002', 'Bラベル');

  const session = createPanelArrangeSession('sheet-1', panels);

  assert.equal(session.tokens.length, 2);
  assert.equal(session.tokens[0].assignedPanelIndex, 0);
  assert.equal(session.tokens[0].content.originalName, 'E1001.png');
  assert.equal(session.tokens[0].content.freeLabels[0].text, 'Aラベル');
  assert.equal(isPanelArrangeSessionComplete(session), true);
});

test('dropping onto an occupied panel leaves the previous image unresolved instead of swapping', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel(panels[0], 'image-a', 'E1001', 'Aラベル');
  panels[1] = buildImagePanel(panels[1], 'image-b', 'E1002', 'Bラベル');
  const initial = createPanelArrangeSession('sheet-1', panels);

  const result = stagePanelArrangeDrop(initial, initial.tokens[0].id, 1, panels);
  const view = buildPanelArrangeView(panels, result.session);

  assert.equal(result.status, 'placed');
  assert.equal(getUnresolvedPanelArrangeTokens(result.session).length, 1);
  assert.equal(view.panels[0].imageId, null);
  assert.equal(view.panels[1].imageId, 'image-a');
  assert.equal(view.floatingTokensByPanel[1][0].content.imageId, 'image-b');
  assert.equal(view.placedPanelIndices.has(1), true);
  assert.equal(buildPanelArrangeFinalPanels(panels, result.session), null);
});

test('unresolved image can be freely placed before finalizing with labels and layouts intact', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel({ ...panels[0], rowSpan: 2 }, 'image-a', 'E1001', 'Aラベル');
  panels[1] = buildImagePanel({ ...panels[1], colSpan: 2 }, 'image-b', 'E1002', 'Bラベル');
  const initial = createPanelArrangeSession('sheet-1', panels);
  const first = stagePanelArrangeDrop(initial, initial.tokens[0].id, 1, panels).session;
  const unresolved = getUnresolvedPanelArrangeTokens(first)[0];
  const completed = stagePanelArrangeDrop(first, unresolved.id, 0, panels).session;
  const finalPanels = buildPanelArrangeFinalPanels(panels, completed);

  assert.equal(isPanelArrangeSessionComplete(completed), true);
  assert.equal(finalPanels[0].imageId, 'image-b');
  assert.equal(finalPanels[0].rowSpan, 2);
  assert.equal(finalPanels[0].freeLabels[0].text, 'Bラベル');
  assert.equal(finalPanels[1].imageId, 'image-a');
  assert.equal(finalPanels[1].colSpan, 2);
  assert.equal(finalPanels[1].freeLabels[0].text, 'Aラベル');
});

test('merging panels makes tokens from hidden cells unresolved without losing them', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel(panels[0], 'image-a', 'E1001', 'Aラベル');
  panels[1] = buildImagePanel(panels[1], 'image-b', 'E1002', 'Bラベル');
  const session = createPanelArrangeSession('sheet-1', panels);
  const mergedPanels = [...panels];
  mergedPanels[0] = { ...mergedPanels[0], rowSpan: 1, colSpan: 2 };
  mergedPanels[1] = { ...mergedPanels[1], hidden: true };

  const reconciled = reconcilePanelArrangeSession(session, mergedPanels);
  const unresolved = getUnresolvedPanelArrangeTokens(reconciled);

  assert.equal(unresolved.length, 1);
  assert.equal(unresolved[0].content.imageId, 'image-b');
  assert.equal(unresolved[0].floatingPanelIndex, 0);
});

test('non-image dummy content is protected from arrange-mode overwrite', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel(panels[0], 'image-a', 'E1001', 'Aラベル');
  panels[1] = { ...panels[1], label: 'タイトル' };
  const session = createPanelArrangeSession('sheet-1', panels);

  const result = stagePanelArrangeDrop(session, session.tokens[0].id, 1, panels);

  assert.equal(result.status, 'blocked-content');
  assert.equal(result.session, session);
});

test('dummy panel remains a dummy even when legacy image data is still present', () => {
  const panels = buildDefaultPanels();
  panels[0] = buildImagePanel(panels[0], 'image-a', 'E1001', 'Aラベル');
  panels[1] = {
    ...panels[1],
    image: 'data:image/png;base64,legacy',
    label: '埋草',
    code: 'ダミーコマ'
  };

  const session = createPanelArrangeSession('sheet-1', panels);
  const result = stagePanelArrangeDrop(session, session.tokens[0].id, 1, panels);
  const view = buildPanelArrangeView(panels, session);

  assert.equal(session.tokens.length, 1);
  assert.equal(result.status, 'blocked-content');
  assert.equal(view.panels[1].label, '埋草');
  assert.equal(view.panels[1].image, 'data:image/png;base64,legacy');
});
