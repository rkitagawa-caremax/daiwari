const hashString = (value = '') => {
  let hash = 0;
  for (let index = 0; index < value.length; index++) {
    hash = ((hash << 5) - hash + value.charCodeAt(index)) | 0;
  }
  return Math.abs(hash).toString(36);
};

export const normalizeStockImageEntry = (item, imageDataById = {}) => {
  if (!item) return null;
  const resolvedData = item.data || item.image || (item.imageId ? imageDataById[item.imageId] : null);
  if (!resolvedData) return null;

  const stableId = item.id || item.imageId || `legacy-${hashString(resolvedData)}`;
  return {
    id: stableId,
    name: item.name || item.originalName || item.code || `stock-${stableId}.png`,
    data: resolvedData,
    createdAt: item.createdAt || { seconds: Date.now() / 1000 }
  };
};

export const normalizeStockImages = (items = [], imageDataById = {}) => {
  const normalized = [];
  const seenIds = new Set();
  const seenData = new Set();

  items.forEach((item) => {
    const next = normalizeStockImageEntry(item, imageDataById);
    if (!next) return;
    if (seenIds.has(next.id) || seenData.has(next.data)) return;
    seenIds.add(next.id);
    seenData.add(next.data);
    normalized.push(next);
  });

  return normalized;
};

export const isSameStockImageList = (leftItems = [], rightItems = []) => {
  if (leftItems.length !== rightItems.length) return false;
  for (let index = 0; index < leftItems.length; index++) {
    const left = leftItems[index];
    const right = rightItems[index];
    if ((left?.id || null) !== (right?.id || null)) return false;
    if ((left?.data || null) !== (right?.data || null)) return false;
    if ((left?.name || null) !== (right?.name || null)) return false;
  }
  return true;
};
