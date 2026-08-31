import {
  applyPanelTransferableContent,
  clearPanelTransferableContent,
  getPanelTransferableContent,
  hasPanelTransferableContent
} from './panels.js';

export const hasPanelImageContent = (panel = {}) => (
  !!(panel.image || panel.imageId)
  && !panel.label
  && !panel.isText
);

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

export const createPanelArrangeSessionForSheets = (sheetEntries = []) => {
  const normalizedEntries = sheetEntries
    .map((entry) => ({
      sheetId: entry?.sheetId || entry?.id || null,
      panels: entry?.panels || []
    }))
    .filter((entry) => entry.sheetId);
  const sheetIds = normalizedEntries.map((entry) => entry.sheetId);

  return {
    sheetId: sheetIds[0] || null,
    sheetIds,
    tokens: normalizedEntries.flatMap(({ sheetId, panels }) => (
      panels.flatMap((panel, panelIndex) => (
        hasPanelImageContent(panel)
          ? [{
            id: createTokenId(sheetId, panelIndex, panel),
            content: getArrangeContent(panel),
            originalSheetId: sheetId,
            originalPanelIndex: panelIndex,
            assignedSheetId: sheetId,
            assignedPanelIndex: panelIndex,
            floatingSheetId: sheetId,
            floatingPanelIndex: panelIndex,
            isPlaced: false
          }]
          : []
      ))
    ))
  };
};

export const createPanelArrangeSession = (sheetId, panels = []) => (
  createPanelArrangeSessionForSheets([{ sheetId, panels }])
);

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

const resolveTokenSheetId = (session, token, field) => (
  token?.[field]
  || token?.originalSheetId
  || session?.sheetId
  || null
);

export const getPanelArrangeSessionSheetIds = (session) => {
  if (!session) return [];
  const ids = Array.isArray(session.sheetIds) && session.sheetIds.length > 0
    ? session.sheetIds
    : [session.sheetId];
  return ids.filter((id, index) => id && ids.indexOf(id) === index);
};

export const reconcilePanelArrangeSessionForSheets = (session, panelsBySheetId = {}) => {
  if (!session) return null;
  const occupied = new Set();

  return {
    ...session,
    sheetIds: getPanelArrangeSessionSheetIds(session),
    tokens: session.tokens.map((token) => {
      const assignedSheetId = Number.isInteger(token.assignedPanelIndex)
        ? resolveTokenSheetId(session, token, 'assignedSheetId')
        : null;
      const assignedPanels = panelsBySheetId[assignedSheetId] || [];
      const assignmentKey = `${assignedSheetId}:${token.assignedPanelIndex}`;
      const assignedIsVisible = !!assignedSheetId
        && isVisiblePanelIndex(assignedPanels, token.assignedPanelIndex);
      const hasDuplicateAssignment = assignedIsVisible && occupied.has(assignmentKey);
      if (assignedIsVisible && !hasDuplicateAssignment) {
        occupied.add(assignmentKey);
        return {
          ...token,
          originalSheetId: resolveTokenSheetId(session, token, 'originalSheetId'),
          assignedSheetId,
          floatingSheetId: resolveTokenSheetId(session, token, 'floatingSheetId')
        };
      }

      const requestedFloatingSheetId = resolveTokenSheetId(session, token, 'floatingSheetId');
      const originalSheetId = resolveTokenSheetId(session, token, 'originalSheetId');
      const floatingSheetId = panelsBySheetId[requestedFloatingSheetId]
        ? requestedFloatingSheetId
        : panelsBySheetId[originalSheetId]
          ? originalSheetId
          : getPanelArrangeSessionSheetIds(session).find((sheetId) => panelsBySheetId[sheetId]) || null;
      const floatingPanels = panelsBySheetId[floatingSheetId] || [];
      const anchorSource = Number.isInteger(token.floatingPanelIndex)
        ? token.floatingPanelIndex
        : token.originalPanelIndex;
      const floatingPanelIndex = isVisiblePanelIndex(floatingPanels, anchorSource)
        ? anchorSource
        : findVisiblePanelCoveringIndex(floatingPanels, anchorSource);
      return {
        ...token,
        originalSheetId,
        assignedSheetId: null,
        assignedPanelIndex: null,
        floatingSheetId,
        floatingPanelIndex,
        isPlaced: false
      };
    })
  };
};

export const reconcilePanelArrangeSession = (session, panels = [], sheetId = session?.sheetId) => {
  if (!session || !sheetId) return session || null;
  const relevantTokenIds = new Set(session.tokens
    .filter((token) => {
      const assignedSheetId = Number.isInteger(token.assignedPanelIndex)
        ? resolveTokenSheetId(session, token, 'assignedSheetId')
        : null;
      const floatingSheetId = resolveTokenSheetId(session, token, 'floatingSheetId');
      return assignedSheetId === sheetId || floatingSheetId === sheetId;
    })
    .map((token) => token.id));
  const partialSession = { ...session, tokens: session.tokens.filter((token) => relevantTokenIds.has(token.id)) };
  const reconciled = reconcilePanelArrangeSessionForSheets(partialSession, { [sheetId]: panels });
  const reconciledById = new Map(reconciled.tokens.map((token) => [token.id, token]));
  return {
    ...session,
    sheetIds: getPanelArrangeSessionSheetIds(session),
    tokens: session.tokens.map((token) => reconciledById.get(token.id) || token)
  };
};

export const getUnresolvedPanelArrangeTokens = (session) => (
  session?.tokens?.filter((token) => (
    !token.assignedSheetId || !Number.isInteger(token.assignedPanelIndex)
  )) || []
);

export const isPanelArrangeSessionComplete = (session) => (
  !!session && getUnresolvedPanelArrangeTokens(session).length === 0
);

export const stagePanelArrangeDropAcrossSheets = (
  session,
  tokenId,
  targetSheetId,
  targetPanelIndex,
  panelsBySheetId = {}
) => {
  const panels = panelsBySheetId[targetSheetId] || [];
  if (
    !session
    || !tokenId
    || !getPanelArrangeSessionSheetIds(session).includes(targetSheetId)
    || !isVisiblePanelIndex(panels, targetPanelIndex)
  ) {
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
    token.id !== tokenId
    && resolveTokenSheetId(session, token, 'assignedSheetId') === targetSheetId
    && token.assignedPanelIndex === targetPanelIndex
  ));

  const nextSession = {
    ...session,
    tokens: session.tokens.map((token) => {
      if (token.id === tokenId) {
        return {
          ...token,
          assignedSheetId: targetSheetId,
          assignedPanelIndex: targetPanelIndex,
          floatingSheetId: targetSheetId,
          floatingPanelIndex: targetPanelIndex,
          isPlaced: true
        };
      }
      if (token.id === displacedToken?.id) {
        return {
          ...token,
          assignedSheetId: null,
          assignedPanelIndex: null,
          floatingSheetId: targetSheetId,
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

export const stagePanelArrangeDrop = (session, tokenId, targetPanelIndex, panels = []) => (
  stagePanelArrangeDropAcrossSheets(
    session,
    tokenId,
    session?.sheetId,
    targetPanelIndex,
    { [session?.sheetId]: panels }
  )
);

export const buildPanelArrangeViews = (panelsBySheetId = {}, session) => {
  const reconciledSession = reconcilePanelArrangeSessionForSheets(session, panelsBySheetId);
  const viewsBySheetId = {};

  Object.entries(panelsBySheetId).forEach(([sheetId, panels]) => {
    viewsBySheetId[sheetId] = {
      panels: panels.map((panel) => (
        hasPanelImageContent(panel) ? clearArrangeContent(panel) : { ...panel }
      )),
      assignedTokenIdsByPanel: {},
      placedPanelIndices: new Set(),
      floatingTokensByPanel: {}
    };
  });

  reconciledSession?.tokens?.forEach((token) => {
    if (token.assignedSheetId && Number.isInteger(token.assignedPanelIndex)) {
      const view = viewsBySheetId[token.assignedSheetId];
      if (!view) return;
      const targetPanel = view.panels[token.assignedPanelIndex] || {};
      view.panels[token.assignedPanelIndex] = applyArrangeContent(targetPanel, token.content);
      view.assignedTokenIdsByPanel[token.assignedPanelIndex] = token.id;
      if (token.isPlaced) view.placedPanelIndices.add(token.assignedPanelIndex);
      return;
    }

    if (!token.floatingSheetId || !Number.isInteger(token.floatingPanelIndex)) return;
    const view = viewsBySheetId[token.floatingSheetId];
    if (!view) return;
    if (!view.floatingTokensByPanel[token.floatingPanelIndex]) {
      view.floatingTokensByPanel[token.floatingPanelIndex] = [];
    }
    view.floatingTokensByPanel[token.floatingPanelIndex].push(token);
  });

  return { session: reconciledSession, viewsBySheetId };
};

export const buildPanelArrangeView = (panels = [], session) => {
  const workspaceView = buildPanelArrangeViews({ [session?.sheetId]: panels }, session);
  const sheetView = workspaceView.viewsBySheetId[session?.sheetId] || {
    panels,
    assignedTokenIdsByPanel: {},
    placedPanelIndices: new Set(),
    floatingTokensByPanel: {}
  };

  return {
    ...sheetView,
    session: workspaceView.session
  };
};

export const buildPanelArrangeFinalPanelsForSheets = (panelsBySheetId = {}, session) => {
  const workspaceView = buildPanelArrangeViews(panelsBySheetId, session);
  if (!isPanelArrangeSessionComplete(workspaceView.session)) return null;

  return Object.fromEntries(Object.entries(workspaceView.viewsBySheetId).map(([sheetId, view]) => [
    sheetId,
    view.panels.map((panel) => (panel.hidden ? clearArrangeContent(panel) : panel))
  ]));
};

export const buildPanelArrangeFinalPanels = (panels = [], session) => {
  const finalPanelsBySheetId = buildPanelArrangeFinalPanelsForSheets(
    { [session?.sheetId]: panels },
    session
  );
  return finalPanelsBySheetId?.[session?.sheetId] || null;
};
