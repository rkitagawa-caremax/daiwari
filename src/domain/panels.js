export const PANEL_COUNT = 16;

export const DEFAULT_PANEL_DATA = {
  image: null,
  imageId: null,
  text: '',
  label: null,
  code: null,
  rowSpan: 1,
  colSpan: 1,
  hidden: false,
  sizeType: '1/16（1コマ）',
  isText: false
};

const normalizeComparableValue = (value) => (value === undefined ? null : value);

export const isPanelDataEqual = (a, b) => {
  if (a === b) return true;
  if (!a || !b) return false;

  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete('fromTempId');
  keys.delete('fromExcludedId');

  for (const key of keys) {
    const leftValue = normalizeComparableValue(a[key]);
    const rightValue = normalizeComparableValue(b[key]);
    const isObjectLike = (value) => value !== null && typeof value === 'object';

    if (isObjectLike(leftValue) || isObjectLike(rightValue)) {
      if (JSON.stringify(leftValue) !== JSON.stringify(rightValue)) return false;
      continue;
    }
    if (leftValue !== rightValue) return false;
  }
  return true;
};

export const getPanelDataPatch = (currentPanel = {}, nextPanel = {}) => {
  const patch = {};
  const keys = new Set([...Object.keys(currentPanel), ...Object.keys(nextPanel)]);

  keys.forEach((key) => {
    const currentValue = normalizeComparableValue(currentPanel[key]);
    const nextValue = normalizeComparableValue(nextPanel[key]);
    const isObjectLike = (value) => value !== null && typeof value === 'object';
    const isChanged = (isObjectLike(currentValue) || isObjectLike(nextValue))
      ? JSON.stringify(currentValue) !== JSON.stringify(nextValue)
      : currentValue !== nextValue;

    if (isChanged) {
      patch[key] = nextPanel[key] === undefined ? null : nextPanel[key];
    }
  });

  return patch;
};

export const hasPanelTransferableContent = (panel = {}) => {
  return !!(panel.image || panel.imageId || panel.label || panel.isText || panel.code);
};

export const clearPanelTransferableContent = (panel = {}) => ({
  ...panel,
  image: null,
  imageId: null,
  label: null,
  code: null,
  text: '',
  isText: false
});

export const sanitizePanelData = (panel = {}) => {
  const sanitized = {};
  Object.keys(panel || {}).forEach((key) => {
    sanitized[key] = panel[key] === undefined ? null : panel[key];
  });
  if (sanitized.imageId) {
    sanitized.image = null;
  }
  return sanitized;
};

export const buildDefaultPanels = () => (
  Array.from({ length: PANEL_COUNT }, () => ({ ...DEFAULT_PANEL_DATA }))
);

export const getPanelsFromDocData = (data = {}) => {
  const basePanels = buildDefaultPanels();
  const fallbackPanels = Array.isArray(data.panels) ? data.panels : [];
  for (let index = 0; index < PANEL_COUNT; index++) {
    basePanels[index] = {
      ...basePanels[index],
      ...(fallbackPanels[index] || {})
    };
  }

  const mapPanels = (data.panelsMap && typeof data.panelsMap === 'object')
    ? data.panelsMap
    : null;
  if (!mapPanels) return basePanels;

  Object.entries(mapPanels).forEach(([rawIndex, panel]) => {
    const index = Number.parseInt(rawIndex, 10);
    if (Number.isNaN(index) || index < 0 || index >= PANEL_COUNT) return;
    if (!panel || typeof panel !== 'object' || Array.isArray(panel)) return;
    basePanels[index] = {
      ...basePanels[index],
      ...panel
    };
  });

  return basePanels;
};

export const toPanelsMap = (panels = []) => {
  const map = {};
  for (let index = 0; index < PANEL_COUNT; index++) {
    map[String(index)] = sanitizePanelData({
      ...DEFAULT_PANEL_DATA,
      ...(panels[index] || {})
    });
  }
  return map;
};

export const buildPanelMapUpdates = (prevPanels = [], nextPanels = []) => {
  const updates = {};
  for (let index = 0; index < PANEL_COUNT; index++) {
    const previousPanel = { ...DEFAULT_PANEL_DATA, ...(prevPanels[index] || {}) };
    const nextPanel = { ...DEFAULT_PANEL_DATA, ...(nextPanels[index] || {}) };
    if (!isPanelDataEqual(previousPanel, nextPanel)) {
      updates[`panelsMap.${index}`] = sanitizePanelData(nextPanel);
    }
  }
  return updates;
};
