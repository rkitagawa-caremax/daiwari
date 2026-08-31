import {
  DAIWARI_PANEL_DROPZONE_PREFIX,
  extractPanelArrangeDragPayload
} from './dragPayload.js';

export const parsePanelDropZone = (zoneId) => {
  if (!zoneId?.startsWith(DAIWARI_PANEL_DROPZONE_PREFIX)) return null;

  const panelTarget = zoneId.slice(DAIWARI_PANEL_DROPZONE_PREFIX.length);
  const separatorIndex = panelTarget.lastIndexOf(':');
  if (separatorIndex === -1) return null;

  const sheetId = panelTarget.slice(0, separatorIndex);
  const panelIndex = Number.parseInt(panelTarget.slice(separatorIndex + 1), 10);
  if (!sheetId || Number.isNaN(panelIndex)) return null;

  return { sheetId, panelIndex };
};

export const dispatchWorkspaceDrop = ({
  zoneId,
  dragPayload = {},
  onDropToPanel,
  onDropToTemp,
  onDropToStock,
  onDropToExcluded
}) => {
  if (!zoneId) return false;

  const isPanelDropZone = zoneId.startsWith(DAIWARI_PANEL_DROPZONE_PREFIX);
  if (extractPanelArrangeDragPayload(dragPayload) && !isPanelDropZone) {
    return false;
  }

  if (zoneId === 'temp') return onDropToTemp(dragPayload);
  if (zoneId === 'stock') return onDropToStock(dragPayload);
  if (zoneId === 'excluded') return onDropToExcluded(dragPayload);

  const panelTarget = parsePanelDropZone(zoneId);
  if (!panelTarget) return false;
  return onDropToPanel(panelTarget.sheetId, panelTarget.panelIndex, dragPayload);
};
