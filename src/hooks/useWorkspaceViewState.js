import { useMemo } from 'react';

import {
  buildPanelArrangeViews,
  getPanelArrangeSessionSheetIds,
  getUnresolvedPanelArrangeTokens
} from '../domain/panelArrange';
import { normalizeCode } from '../domain/productCodes';
import {
  getAdjacentPageOptions,
  getTwoPageDisplaySheets
} from '../domain/twoPageWorkspace';

const getVisibleSheets = (sheets, genreFilter) => (
  genreFilter === 'all'
    ? sheets
    : sheets.filter((sheet) => sheet.genre === genreFilter)
);

export const useWorkspaceViewState = ({
  sheets,
  images,
  viewMode,
  genreFilter,
  activeSheetId,
  secondarySheetId,
  panelArrangeSession
}) => {
  const filteredSheets = useMemo(
    () => getVisibleSheets(sheets, genreFilter),
    [genreFilter, sheets]
  );

  const displaySheets = useMemo(() => (
    viewMode === 'single'
      ? getTwoPageDisplaySheets(sheets, activeSheetId, secondarySheetId)
      : filteredSheets
  ), [activeSheetId, filteredSheets, secondarySheetId, sheets, viewMode]);

  const adjacentPageOptions = useMemo(() => (
    viewMode === 'single'
      ? getAdjacentPageOptions(sheets, activeSheetId)
      : []
  ), [activeSheetId, sheets, viewMode]);

  const isTwoPageMode = viewMode === 'single'
    && !!secondarySheetId
    && displaySheets.some((sheet) => sheet.id === secondarySheetId);

  const panelArrangeWorkspaceView = useMemo(() => {
    if (!panelArrangeSession) return null;

    const panelsBySheetId = Object.fromEntries(
      getPanelArrangeSessionSheetIds(panelArrangeSession).map((sheetId) => {
        const targetSheet = sheets.find((sheet) => sheet.id === sheetId);
        return [sheetId, targetSheet?.panels || []];
      })
    );

    return buildPanelArrangeViews(panelsBySheetId, panelArrangeSession);
  }, [panelArrangeSession, sheets]);

  const unresolvedPanelArrangeCount = useMemo(() => (
    getUnresolvedPanelArrangeTokens(panelArrangeWorkspaceView?.session || panelArrangeSession).length
  ), [panelArrangeSession, panelArrangeWorkspaceView]);

  const imageDataById = useMemo(() => Object.fromEntries(
    images
      .filter((image) => image?.id && image?.data)
      .map((image) => [image.id, image.data])
  ), [images]);

  const currentList = viewMode === 'single' ? sheets : filteredSheets;

  const currentIndex = useMemo(
    () => currentList.findIndex((sheet) => sheet.id === activeSheetId),
    [activeSheetId, currentList]
  );

  const salesLookupVisibleCodes = useMemo(() => {
    if (viewMode !== 'single' || !activeSheetId) return null;
    const activeSheet = sheets.find((sheet) => sheet.id === activeSheetId);
    if (!activeSheet?.panels) return [];

    const codes = new Set();
    activeSheet.panels.forEach((panel) => {
      if (!panel || panel.hidden) return;
      const normalized = normalizeCode(panel.code || '');
      if (normalized) codes.add(normalized);
    });
    return Array.from(codes);
  }, [activeSheetId, sheets, viewMode]);

  const activeSheetLabelCount = useMemo(() => {
    if (viewMode !== 'single' || !activeSheetId) return 0;
    const targetSheet = sheets.find((sheet) => sheet.id === activeSheetId);
    if (!targetSheet?.panels) return 0;

    return targetSheet.panels.reduce((count, panel) => {
      const freeLabelsCount = panel?.freeLabels?.length || 0;
      const hasLegacyLabel = !!panel?.freeText && freeLabelsCount === 0;
      return count + freeLabelsCount + (hasLegacyLabel ? 1 : 0);
    }, 0);
  }, [activeSheetId, sheets, viewMode]);

  return {
    displaySheets,
    adjacentPageOptions,
    isTwoPageMode,
    panelArrangeWorkspaceView,
    unresolvedPanelArrangeCount,
    imageDataById,
    currentList,
    currentIndex,
    salesLookupVisibleCodes,
    activeSheetLabelCount
  };
};
