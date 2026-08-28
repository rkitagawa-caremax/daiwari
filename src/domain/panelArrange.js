import {
  applyPanelTransferableContent,
  clearPanelTransferableContent,
  getPanelTransferableContent,
  hasPanelTransferableContent
} from './panels.js';

export const hasPanelImageContent = (panel = {}) => !!(panel.image || panel.imageId);

const getArrangeContent = (panel = {}) => ({
  ...getPanelTransferableContent(panel),
  originalName: panel.originalName || null
});

const clearArrangeContent = (panel = {}) => ({
  ...clearPanelTransferableContent(panel),
  originalName: null
});

const applyArrangeContent = (panel = {}, content = {}) => ({
  ...applyPanelTransferableContent(panel, content),
  originalName: content.originalName || null
});

const createTokenId = (sheetId, panelIndex, panel = {}) => (
  `${sheetId}:${panelIndex}:${panel.imageId || panel.code || 'image'}`
);

export const createPanelArrangeSession = (sheetId, panels = []) => ({
  sheetId,
  tokens: panels.flatMap((panel, panelIndex) => (
    hasPanelImageContent(panel)
      ? [{
        id: createTokenId(sheetId, panelIndex, panel),
        content: getArrangeContent(panel),
        originalPanelIndex: panelIndex,
        assignedPanelIndex: panelIndex,
        floatingPanelIndex: panelIndex,
        isPlaced: false
      }]
      : []
  ))
});

const isVisiblePanelIndex = (panels, panelIndex) => (
  Number.isInteger(panelIndex)
  && panelIndex >= 0
  && panelIndex < panels.length
  && !panels[panelIndex]?.hidden
);

const findVisiblePanelCoveringIndex = (panels = [], targetIndex) => {
  const targetRow = Math.floor(targetIndex / 4);
  const targetCol = targetIndex % 4;

  for (let panelIndex = 0; panelIndex < panels.length; panelIndex += 1) {
    const panel = panels[panelIndex] || {};
    if (panel.hidden) continue;
    const startRow = Math.floor(panelIndex / 4);
    const startCol = panelIndex % 4;
    const rowSpan = panel.rowSpan || 1;
    const colSpan = panel.colSpan || 1;
    if (
      targetRow >= startRow
      && targetRow < startRow + rowSpan
      && targetCol >= startCol
      && targetCol < startCol + colSpan
    ) return panelIndex;
  }

  return panels.findIndex((panel) => !panel?.hidden);
};

export const reconcilePanelArrangeSession = (session, panels = []) => {
  if (!session) return null;
  const occupied = new Set();

  return {
    ...session,
    tokens: session.tokens.map((token) => {
      const assignedIsVisible = isVisiblePanelIndex(panels, token.assignedPanelIndex);
      const hasDuplicateAssignment = assignedIsVisible && occupied.has(token.assignedPanelIndex);
      if (assignedIsVisible && !hasDuplicateAssignment) {
        occupied.add(token.assignedPanelIndex);
        return token;
      }

      const anchorSource = Number.isInteger(token.floatingPanelIndex)
        ? token.floatingPanelIndex
        : token.originalPanelIndex;
      const floatingPanelIndex = isVisiblePanelIndex(panels, anchorSource)
        ? anchorSource
        : findVisiblePanelCoveringIndex(panels, anchorSource);
      return {
        ...token,
        assignedPanelIndex: null,
        floatingPanelIndex,
        isPlaced: false
      };
    })
  };
};

export const getUnresolvedPanelArrangeTokens = (session) => (
  session?.tokens?.filter((token) => !Number.isInteger(token.assignedPanelIndex)) || []
);

export const isPanelArrangeSessionComplete = (session) => (
  !!session && getUnresolvedPanelArrangeTokens(session).length === 0
);

export const stagePanelArrangeDrop = (session, tokenId, targetPanelIndex, panels = []) => {
  if (!session || !tokenId || !isVisiblePanelIndex(panels, targetPanelIndex)) {
    return { status: 'invalid', session };
  }

  const sourceToken = session.tokens.find((token) => token.id === tokenId);
  if (!sourceToken) return { status: 'invalid', session };

  const targetPanel = panels[targetPanelIndex] || {};
  const hasBlockingNonImageContent = (
    hasPanelTransferableContent(targetPanel)
    && !hasPanelImageContent(targetPanel)
  );
  if (hasBlockingNonImageContent) {
    return { status: 'blocked-content', session };
  }

  const displacedToken = session.tokens.find((token) => (
    token.id !== tokenId && token.assignedPanelIndex === targetPanelIndex
  ));

  const nextSession = {
    ...session,
    tokens: session.tokens.map((token) => {
      if (token.id === tokenId) {
        return {
          ...token,
          assignedPanelIndex: targetPanelIndex,
          floatingPanelIndex: targetPanelIndex,
          isPlaced: true
        };
      }
      if (token.id === displacedToken?.id) {
        return {
          ...token,
          assignedPanelIndex: null,
          floatingPanelIndex: targetPanelIndex,
          isPlaced: false
        };
      }
      return token;
    })
  };

  return {
    status: 'placed',
    session: nextSession,
    displacedTokenId: displacedToken?.id || null
  };
};

export const buildPanelArrangeView = (panels = [], session) => {
  const reconciledSession = reconcilePanelArrangeSession(session, panels);
  const nextPanels = panels.map((panel) => (
    hasPanelImageContent(panel) ? clearArrangeContent(panel) : { ...panel }
  ));
  const assignedTokenIdsByPanel = {};
  const placedPanelIndices = new Set();
  const floatingTokensByPanel = {};

  reconciledSession?.tokens?.forEach((token) => {
    if (Number.isInteger(token.assignedPanelIndex)) {
      const targetPanel = nextPanels[token.assignedPanelIndex] || {};
      nextPanels[token.assignedPanelIndex] = applyArrangeContent(targetPanel, token.content);
      assignedTokenIdsByPanel[token.assignedPanelIndex] = token.id;
      if (token.isPlaced) placedPanelIndices.add(token.assignedPanelIndex);
      return;
    }

    if (!Number.isInteger(token.floatingPanelIndex)) return;
    if (!floatingTokensByPanel[token.floatingPanelIndex]) {
      floatingTokensByPanel[token.floatingPanelIndex] = [];
    }
    floatingTokensByPanel[token.floatingPanelIndex].push(token);
  });

  return {
    panels: nextPanels,
    session: reconciledSession,
    assignedTokenIdsByPanel,
    placedPanelIndices,
    floatingTokensByPanel
  };
};

export const buildPanelArrangeFinalPanels = (panels = [], session) => {
  const view = buildPanelArrangeView(panels, session);
  if (!isPanelArrangeSessionComplete(view.session)) return null;
  return view.panels.map((panel) => (
    panel.hidden ? clearArrangeContent(panel) : panel
  ));
};
