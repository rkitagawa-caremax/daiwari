import { getPanelFreeLabels } from '../domain/panels.js';

const DAIWARI_DRAG_PAYLOAD_TYPE = 'application/x-daiwari-drag';
const DAIWARI_DRAG_PAYLOAD_PREFIX = '__daiwari_drag__:';
export const DAIWARI_DROPZONE_ATTR = 'data-daiwari-dropzone-id';
export const DAIWARI_PANEL_DROPZONE_PREFIX = 'panel:';
export const POINTER_DRAG_THRESHOLD_PX = 10;
const DAIWARI_DROP_HANDLED_FLAG = '__daiwariDropHandled';
let activeNativeDragPayload = null;
let activePanelMovePayload = null;
let isNativeDragSessionActive = false;

const DAIWARI_DRAG_FIELDS = [
  'moveSourceType',
  'sourceSheetId',
  'sourceIndex',
  'textData',
  'type',
  'src',
  'imageId',
  'label',
  'code',
  'name',
  'isText',
  'hasTextPayload',
  'textPayload',
  'text',
  'freeLabels',
  'freeText',
  'fromTempId',
  'fromExcludedId'
];

export const normalizeDragPayload = (payload = {}) => {
  const normalized = { __daiwariDragPayload: true };
  DAIWARI_DRAG_FIELDS.forEach((field) => {
    const value = payload[field];
    if (field === 'freeLabels' && Array.isArray(value)) {
      normalized[field] = JSON.stringify(value);
      return;
    }
    normalized[field] = value === undefined || value === null ? '' : String(value);
  });
  return normalized;
};

export const getActiveNativeDragPayload = () => {
  if (!isNativeDragSessionActive || !activeNativeDragPayload) return null;
  return normalizeDragPayload(activeNativeDragPayload);
};

export const clearActiveNativeDragPayload = () => {
  activeNativeDragPayload = null;
  activePanelMovePayload = null;
  isNativeDragSessionActive = false;
};

export const isDropEventHandled = (nativeEvent) => {
  return !!(nativeEvent && nativeEvent[DAIWARI_DROP_HANDLED_FLAG] === true);
};

export const markDropEventHandled = (nativeEvent) => {
  if (!nativeEvent) return;
  nativeEvent[DAIWARI_DROP_HANDLED_FLAG] = true;
};

export const setDragPayload = (dataTransfer, payload = {}) => {
  if (!dataTransfer) return;

  const normalized = normalizeDragPayload(payload);
  activeNativeDragPayload = normalized;
  activePanelMovePayload = normalized.moveSourceType === 'panel' ? normalized : null;
  isNativeDragSessionActive = true;
  const serialized = `${DAIWARI_DRAG_PAYLOAD_PREFIX}${JSON.stringify(normalized)}`;

  DAIWARI_DRAG_FIELDS.forEach((field) => {
    try {
      dataTransfer.setData(field, normalized[field]);
    } catch (error) {
      void error;
    }
  });

  [DAIWARI_DRAG_PAYLOAD_TYPE, 'text/plain', 'text'].forEach((type) => {
    try {
      dataTransfer.setData(type, serialized);
    } catch (error) {
      void error;
    }
  });
};

const getDragData = (dataTransfer, type) => {
  if (!dataTransfer) return '';
  try {
    return dataTransfer.getData(type) || '';
  } catch {
    return '';
  }
};

const parseDragPayload = (rawValue) => {
  if (!rawValue || !rawValue.startsWith(DAIWARI_DRAG_PAYLOAD_PREFIX)) return null;
  try {
    const parsed = JSON.parse(rawValue.slice(DAIWARI_DRAG_PAYLOAD_PREFIX.length));
    if (!parsed || parsed.__daiwariDragPayload !== true) return null;
    return normalizeDragPayload(parsed);
  } catch (error) {
    void error;
    return null;
  }
};

export const getDragPayload = (dataTransfer) => {
  const parsedPayload = (
    parseDragPayload(getDragData(dataTransfer, DAIWARI_DRAG_PAYLOAD_TYPE))
    || parseDragPayload(getDragData(dataTransfer, 'text/plain'))
    || parseDragPayload(getDragData(dataTransfer, 'text'))
  );
  if (parsedPayload) {
    activeNativeDragPayload = parsedPayload;
    activePanelMovePayload = parsedPayload.moveSourceType === 'panel' ? parsedPayload : null;
    return parsedPayload;
  }

  const legacyPayload = {};
  DAIWARI_DRAG_FIELDS.forEach((field) => {
    const value = getDragData(dataTransfer, field);
    if (value !== '') {
      legacyPayload[field] = value;
    }
  });

  if (Object.keys(legacyPayload).length > 0) {
    const normalizedLegacyPayload = normalizeDragPayload(legacyPayload);
    activeNativeDragPayload = normalizedLegacyPayload;
    activePanelMovePayload = normalizedLegacyPayload.moveSourceType === 'panel' ? normalizedLegacyPayload : null;
    return normalizedLegacyPayload;
  }

  return getActiveNativeDragPayload();
};

export const parseNullableDragValue = (value) => {
  if (value === undefined || value === null) return null;
  const normalized = String(value);
  if (!normalized || normalized === 'null' || normalized === 'undefined') return null;
  return normalized;
};

export const extractPanelMoveDragPayload = (dragPayload = {}) => {
  if ((dragPayload.moveSourceType || '') !== 'panel') return null;
  const sourceSheetId = parseNullableDragValue(dragPayload.sourceSheetId);
  const sourceIndex = Number.parseInt(dragPayload.sourceIndex || '', 10);
  if (!sourceSheetId || Number.isNaN(sourceIndex)) return null;
  return {
    sourceSheetId,
    sourceIndex,
    movedText: dragPayload.textData || ''
  };
};

export const getActivePanelMoveDragPayload = () => {
  if (!isNativeDragSessionActive || !activePanelMovePayload) return null;
  return extractPanelMoveDragPayload(activePanelMovePayload);
};

export const extractPanelAssignmentFromDragPayload = (dragPayload = {}, fallbackText = '') => {
  const src = parseNullableDragValue(dragPayload.src) || '';
  const imageId = parseNullableDragValue(dragPayload.imageId);
  const label = parseNullableDragValue(dragPayload.label);
  let code = parseNullableDragValue(dragPayload.code);
  const fileName = dragPayload.name || '';
  const isText = dragPayload.isText === 'true';
  const hasTextPayload = dragPayload.hasTextPayload === '1';
  const transferredText = hasTextPayload ? (dragPayload.textPayload || '') : (dragPayload.text || '');
  let parsedFreeLabels = [];
  try {
    const parsed = JSON.parse(dragPayload.freeLabels || '[]');
    parsedFreeLabels = Array.isArray(parsed) ? parsed : [];
  } catch {
    parsedFreeLabels = [];
  }
  const freeLabels = getPanelFreeLabels({
    freeLabels: parsedFreeLabels,
    freeText: parseNullableDragValue(dragPayload.freeText)
  });
  const fromTempId = parseNullableDragValue(dragPayload.fromTempId);
  const fromExcludedId = parseNullableDragValue(dragPayload.fromExcludedId);

  if (!code && fileName && !isText) {
    const match = fileName.match(/[A-Za-z]\d{4}/);
    if (match) {
      code = match[0].toUpperCase();
    }
  }

  if (!src && !imageId && !label && !code && !isText && freeLabels.length === 0 && !fromTempId && !fromExcludedId) {
    return null;
  }

  return {
    image: src || null,
    imageId: imageId || null,
    label: label || null,
    code: code || null,
    isText,
    text: isText
      ? ((hasTextPayload || transferredText !== '') ? transferredText : fallbackText)
      : '',
    freeLabels,
    freeText: null,
    fromTempId: fromTempId || null,
    fromExcludedId: fromExcludedId || null
  };
};
