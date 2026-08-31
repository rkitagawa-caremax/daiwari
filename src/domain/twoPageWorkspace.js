export const getAdjacentPageOptions = (sheets = [], activeSheetId = null) => {
  const activeIndex = sheets.findIndex((sheet) => sheet?.id === activeSheetId);
  if (activeIndex === -1) return [];

  return [
    activeIndex > 0
      ? {
          id: sheets[activeIndex - 1].id,
          pageNumber: activeIndex,
          direction: 'prev',
          directionLabel: '前のページ'
        }
      : null,
    activeIndex < sheets.length - 1
      ? {
          id: sheets[activeIndex + 1].id,
          pageNumber: activeIndex + 2,
          direction: 'next',
          directionLabel: '次のページ'
        }
      : null
  ].filter(Boolean);
};

export const getTwoPageDisplaySheets = (
  sheets = [],
  activeSheetId = null,
  secondarySheetId = null
) => {
  if (!activeSheetId) return [];

  const requestedIds = [activeSheetId, secondarySheetId]
    .filter((id, index, ids) => id && ids.indexOf(id) === index);
  const sheetById = new Map(sheets.map((sheet) => [sheet.id, sheet]));

  return requestedIds
    .map((id) => sheetById.get(id))
    .filter(Boolean);
};

export const getPageNavigationSelection = (
  sheets = [],
  activeSheetId = null,
  secondarySheetId = null,
  direction
) => {
  const activeIndex = sheets.findIndex((sheet) => sheet?.id === activeSheetId);
  if (activeIndex === -1) return null;

  const delta = direction === 'prev' ? -1 : direction === 'next' ? 1 : 0;
  const targetIndex = activeIndex + delta;
  if (delta === 0 || targetIndex < 0 || targetIndex >= sheets.length) return null;

  let nextSecondarySheetId = null;
  const secondaryIndex = sheets.findIndex((sheet) => sheet?.id === secondarySheetId);
  if (secondaryIndex !== -1) {
    const relativeSide = secondaryIndex - activeIndex;
    nextSecondarySheetId = (sheets[targetIndex + relativeSide] || sheets[activeIndex]).id;
  }

  return {
    activeSheetId: sheets[targetIndex].id,
    secondarySheetId: nextSecondarySheetId
  };
};
