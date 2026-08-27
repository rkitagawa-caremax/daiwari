const normalizeSearchText = (value) => String(value || '').trim().toLowerCase();

const matchesSearch = (query, ...values) => (
  values.some((value) => normalizeSearchText(value).includes(query))
);

const createImageKeyStore = () => ({ ids: new Set(), data: new Set() });

const addImageKeys = (target, item) => {
  if (!item) return;
  if (item.id) target.ids.add(item.id);
  if (item.imageId) target.ids.add(item.imageId);
  if (item.data) target.data.add(item.data);
  if (item.image) target.data.add(item.image);
};

const hasImageKey = (keys, item) => {
  if (!item) return false;
  return (
    (item.id && keys.ids.has(item.id))
    || (item.imageId && keys.ids.has(item.imageId))
    || (item.data && keys.data.has(item.data))
    || (item.image && keys.data.has(item.image))
  );
};

/**
 * ライブラリの通常表示を維持しながら、検索中だけ配置済み画像を結果へ合成する。
 * 配置情報は UI 側でドラッグ・削除対象から除外し、ページ遷移にだけ利用する。
 */
export const buildSidebarImageResults = ({
  images = [],
  sheets = [],
  excludedItems = [],
  searchQuery = ''
} = {}) => {
  const safeImages = Array.isArray(images) ? images.filter(Boolean) : [];
  const safeSheets = Array.isArray(sheets) ? sheets.filter(Boolean) : [];
  const query = normalizeSearchText(searchQuery);
  const imagesById = new Map();
  const imagesByData = new Map();

  safeImages.forEach((image) => {
    if (image.id) imagesById.set(image.id, image);
    if (image.data) imagesByData.set(image.data, image);
  });

  const usedImageKeys = createImageKeyStore();
  const assignedMatches = [];
  const seenAssignments = new Set();

  safeSheets.forEach((sheet, sheetIndex) => {
    const panels = Array.isArray(sheet.panels) ? sheet.panels : [];
    panels.forEach((panel, panelIndex) => {
      if (!panel || panel.hidden || (!panel.image && !panel.imageId)) return;

      const stockImage = (
        (panel.imageId ? imagesById.get(panel.imageId) : null)
        || (panel.image ? imagesByData.get(panel.image) : null)
        || null
      );
      addImageKeys(usedImageKeys, panel);
      addImageKeys(usedImageKeys, stockImage);

      if (!query) return;

      const panelCode = panel.code || '';
      const stockCode = stockImage?.code || '';
      const imageName = stockImage?.name || panel.originalName || '';
      if (!matchesSearch(query, panelCode, stockCode, imageName)) return;

      const resolvedData = panel.image || stockImage?.data || null;
      if (!resolvedData) return;

      const imageIdentity = panel.imageId || stockImage?.id || `panel-${panelIndex}`;
      const assignmentKey = `${sheet.id || sheetIndex}:${imageIdentity}`;
      if (seenAssignments.has(assignmentKey)) return;
      seenAssignments.add(assignmentKey);

      assignedMatches.push({
        ...(stockImage || {}),
        id: stockImage?.id || panel.imageId || null,
        data: resolvedData,
        name: imageName || panelCode || stockCode || '配置済み画像',
        code: panelCode || stockCode || null,
        freeLabels: panel.freeLabels || stockImage?.freeLabels || [],
        freeText: panel.freeText || stockImage?.freeText || null,
        searchResultKey: `assigned:${sheet.id || sheetIndex}:${panelIndex}:${imageIdentity}`,
        assignment: {
          sheetId: sheet.id,
          sheetNumber: sheetIndex + 1,
          genre: sheet.genre || 'none',
          panelIndex,
          code: panelCode || stockCode || ''
        }
      });
    });
  });

  const excludedImageKeys = createImageKeyStore();
  (Array.isArray(excludedItems) ? excludedItems : []).forEach((item) => {
    addImageKeys(excludedImageKeys, item);
  });

  let availableImages = safeImages.filter((image) => (
    !hasImageKey(usedImageKeys, image) && !hasImageKey(excludedImageKeys, image)
  ));

  if (query) {
    availableImages = availableImages.filter((image) => (
      matchesSearch(query, image.name, image.code)
    ));
  }

  return query ? [...assignedMatches, ...availableImages] : availableImages;
};
