import { getPanelFreeLabels } from './panels.js';

const cloneCatalogTextData = (value) => {
  if (!value || typeof value !== 'object') return null;
  return {
    ...value,
    itemNumberCandidates: Array.isArray(value.itemNumberCandidates) ? [...value.itemNumberCandidates] : [],
    catchCopyCandidates: Array.isArray(value.catchCopyCandidates) ? [...value.catchCopyCandidates] : [],
    availabilityLabels: Array.isArray(value.availabilityLabels) ? [...value.availabilityLabels] : [],
    handlingMarkers: Array.isArray(value.handlingMarkers) ? [...value.handlingMarkers] : [],
    specifications: Array.isArray(value.specifications) ? [...value.specifications] : [],
    compositionDetails: Array.isArray(value.compositionDetails) ? [...value.compositionDetails] : [],
    materialDetails: Array.isArray(value.materialDetails) ? [...value.materialDetails] : []
  };
};

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
    code: item.code || null,
    sourcePdfName: item.sourcePdfName || null,
    sourcePage: item.sourcePage || null,
    pdfPageNumber: item.pdfPageNumber || null,
    sizeType: item.sizeType || null,
    cropRect: item.cropRect ? { ...item.cropRect } : null,
    textExtractionRect: item.textExtractionRect ? { ...item.textExtractionRect } : null,
    sourceText: typeof item.sourceText === 'string' ? item.sourceText : '',
    sourceTextVersion: item.sourceTextVersion || null,
    sourceTextTruncated: item.sourceTextTruncated === true,
    catalogCode: item.catalogCode || null,
    productName: item.productName || item.catalogName || null,
    productNameSource: item.productNameSource || (item.catalogName ? 'csv' : null),
    priceIncludingTax: Number.isFinite(item.priceIncludingTax) ? item.priceIncludingTax : null,
    priceExcludingTax: Number.isFinite(item.priceExcludingTax) ? item.priceExcludingTax : null,
    priceCandidates: Array.isArray(item.priceCandidates)
      ? item.priceCandidates.map((candidate) => ({ ...candidate }))
      : [],
    priceExtractionConfidence: item.priceExtractionConfidence || null,
    catalogTextData: cloneCatalogTextData(item.catalogTextData),
    freeLabels: getPanelFreeLabels(item),
    freeText: null,
    // 作業したアカウント (アップロード / コマから解除) の UID 一覧。未記録の既存画像は null。
    workedBy: Array.isArray(item.workedBy) && item.workedBy.length > 0 ? [...item.workedBy] : null,
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

export const normalizeCloudImageDocuments = (documents = []) => {
  const loadedImages = documents.map((document) => ({
    ...(typeof document?.data === 'function' ? document.data() : {}),
    id: document?.id
  }));
  const imageDataById = {};

  loadedImages.forEach((image) => {
    if (image?.id && (image?.data || image?.image)) {
      imageDataById[image.id] = image.data || image.image;
    }
  });

  return normalizeStockImages(loadedImages, imageDataById);
};

export const isSameStockImageList = (leftItems = [], rightItems = []) => {
  if (leftItems.length !== rightItems.length) return false;
  for (let index = 0; index < leftItems.length; index++) {
    const left = leftItems[index];
    const right = rightItems[index];
    if ((left?.id || null) !== (right?.id || null)) return false;
    if ((left?.data || null) !== (right?.data || null)) return false;
    if ((left?.name || null) !== (right?.name || null)) return false;
    if ((left?.code || null) !== (right?.code || null)) return false;
    if ((left?.sourcePdfName || null) !== (right?.sourcePdfName || null)) return false;
    if ((left?.sourcePage || null) !== (right?.sourcePage || null)) return false;
    if ((left?.pdfPageNumber || null) !== (right?.pdfPageNumber || null)) return false;
    if ((left?.sizeType || null) !== (right?.sizeType || null)) return false;
    if (JSON.stringify(left?.cropRect || null) !== JSON.stringify(right?.cropRect || null)) return false;
    if (JSON.stringify(left?.textExtractionRect || null) !== JSON.stringify(right?.textExtractionRect || null)) return false;
    if ((left?.sourceText || '') !== (right?.sourceText || '')) return false;
    if ((left?.sourceTextVersion || null) !== (right?.sourceTextVersion || null)) return false;
    if ((left?.sourceTextTruncated === true) !== (right?.sourceTextTruncated === true)) return false;
    if ((left?.catalogCode || null) !== (right?.catalogCode || null)) return false;
    if ((left?.productName || left?.catalogName || null) !== (right?.productName || right?.catalogName || null)) return false;
    if ((left?.productNameSource || null) !== (right?.productNameSource || null)) return false;
    if ((left?.priceIncludingTax ?? null) !== (right?.priceIncludingTax ?? null)) return false;
    if ((left?.priceExcludingTax ?? null) !== (right?.priceExcludingTax ?? null)) return false;
    if (JSON.stringify(left?.priceCandidates || []) !== JSON.stringify(right?.priceCandidates || [])) return false;
    if ((left?.priceExtractionConfidence || null) !== (right?.priceExtractionConfidence || null)) return false;
    if (JSON.stringify(left?.catalogTextData || null) !== JSON.stringify(right?.catalogTextData || null)) return false;
    if (JSON.stringify(getPanelFreeLabels(left)) !== JSON.stringify(getPanelFreeLabels(right))) return false;
    if (JSON.stringify(left?.workedBy || null) !== JSON.stringify(right?.workedBy || null)) return false;
  }
  return true;
};

// 作業者フィルタ: workedBy 未記録の既存画像は全員に表示し、記録済みは本人のみに表示する。
export const isImageWorkedByUser = (image, userId) => {
  const workedBy = Array.isArray(image?.workedBy) ? image.workedBy : null;
  if (!workedBy || workedBy.length === 0) return true;
  if (!userId) return true;
  return workedBy.includes(userId);
};
