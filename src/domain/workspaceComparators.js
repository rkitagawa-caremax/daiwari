import { getPanelFreeLabels, isPanelDataEqual } from './panels.js';

export const toComparableSeconds = (value) => {
  if (!value) return 0;
  if (typeof value.seconds === 'number') return value.seconds;
  if (typeof value.toDate === 'function') {
    try {
      return Math.floor(value.toDate().getTime() / 1000);
    } catch {
      return 0;
    }
  }
  if (typeof value === 'number') return Math.floor(value / 1000);
  return 0;
};

export const isSameTransferItemList = (leftItems = [], rightItems = []) => {
  if (leftItems.length !== rightItems.length) return false;
  for (let index = 0; index < leftItems.length; index++) {
    const left = leftItems[index] || {};
    const right = rightItems[index] || {};
    if ((left.id || '') !== (right.id || '')) return false;
    if ((left.image || null) !== (right.image || null)) return false;
    if ((left.imageId || null) !== (right.imageId || null)) return false;
    if ((left.label || null) !== (right.label || null)) return false;
    if (JSON.stringify(getPanelFreeLabels(left)) !== JSON.stringify(getPanelFreeLabels(right))) return false;
    if ((left.code || null) !== (right.code || null)) return false;
    if ((left.text || '') !== (right.text || '')) return false;
    if (!!left.isText !== !!right.isText) return false;
    if ((left.originalName || '') !== (right.originalName || '')) return false;
    if (toComparableSeconds(left.createdAt) !== toComparableSeconds(right.createdAt)) return false;
  }
  return true;
};

export const isSameSheetList = (leftItems = [], rightItems = []) => {
  if (leftItems.length !== rightItems.length) return false;
  for (let index = 0; index < leftItems.length; index++) {
    const left = leftItems[index] || {};
    const right = rightItems[index] || {};
    if ((left.id || '') !== (right.id || '')) return false;
    if ((left.genre || 'none') !== (right.genre || 'none')) return false;
    if ((left.order || 0) !== (right.order || 0)) return false;
    const leftPanels = Array.isArray(left.panels) ? left.panels : [];
    const rightPanels = Array.isArray(right.panels) ? right.panels : [];
    if (leftPanels.length !== rightPanels.length) return false;
    for (let panelIndex = 0; panelIndex < leftPanels.length; panelIndex++) {
      if (!isPanelDataEqual(leftPanels[panelIndex] || {}, rightPanels[panelIndex] || {})) return false;
    }
  }
  return true;
};
