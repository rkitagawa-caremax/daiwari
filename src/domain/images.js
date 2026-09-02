import { getPanelFreeLabels } from './panels.js';
import { normalizeCode } from './productCodes.js';

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

// --- 画像削除の同一性判定 ---
// ライブラリには同じ介援隊コードの画像が複数たまりうる (PDF切り抜きの再取込は既定で追加登録)。
// 同一バイト列の重複は normalizeStockImages が画面から隠すため、見えている 1 件を消しても
// 隠れた同código・同データの複製が残り、台割CSVの流し込み (コード名の完全一致マッチ) で復活する。
// そこで削除時は「id・画像データ・介援隊コード」のいずれかが一致する複製もまとめて対象にする。
// コードは商品コードの形 (英1-2字+数字3-5桁) に見えるものだけ使い、logo.png のような
// 一般名のファイルを巻き込まないようにする。

const PRODUCT_CODE_PATTERN = /^[A-Z]{1,2}\d{3,5}$/;

export const getStockImageCode = (image) => {
  const name = String(image?.name || image?.originalName || '');
  const stem = name.includes('.') ? name.slice(0, name.lastIndexOf('.')) : name;
  // code フィールド → ファイル名の語幹 → カタログ番号の前置き (261- など) を外した語幹、の順に試す
  const candidates = [image?.code, stem, stem.replace(/^\d+[-_\s]*/, '')];
  for (const candidate of candidates) {
    const normalized = normalizeCode(String(candidate || ''));
    if (PRODUCT_CODE_PATTERN.test(normalized)) return normalized;
  }
  return '';
};

// 削除対象 (id 文字列 or {id, data} など) から同一性キーを作る。
// id しか無い対象は images から現物を引いて code / data を補う。
export const buildImageDeletionIdentity = (targets = [], images = []) => {
  const ids = new Set();
  const data = new Set();
  const codes = new Set();
  const imagesById = new Map(images.filter((image) => image?.id).map((image) => [image.id, image]));

  targets.forEach((target) => {
    if (!target) return;
    const entry = typeof target === 'string' ? { id: target } : target;
    const resolved = entry.id && imagesById.has(entry.id) ? { ...imagesById.get(entry.id), ...entry } : entry;
    if (resolved.id) ids.add(resolved.id);
    const resolvedData = resolved.data || resolved.image || null;
    if (resolvedData) data.add(resolvedData);
    const code = getStockImageCode(resolved);
    if (code) codes.add(code);
  });

  return { ids, data, codes };
};

// 一致の種類を返す: 'direct' = id か画像データの一致 / 'code' = 介援隊コードだけの一致 / null = 不一致
export const getImageDeletionMatchKind = (image, identity) => {
  if (!image || !identity) return null;
  if (image.id && identity.ids.has(image.id)) return 'direct';
  const resolvedData = image.data || image.image || null;
  if (resolvedData && identity.data.has(resolvedData)) return 'direct';
  const code = getStockImageCode(image);
  if (code && identity.codes.has(code)) return 'code';
  return null;
};

// ページに配置中の画像キー (imageId / 画像データ)。コード一致だけの複製でも、
// どこかのコマが参照しているものは削除しない (コマの表示が壊れるため)。
export const collectSheetImageKeys = (sheets = []) => {
  const ids = new Set();
  const data = new Set();
  (Array.isArray(sheets) ? sheets : []).forEach((sheet) => {
    (Array.isArray(sheet?.panels) ? sheet.panels : []).forEach((panel) => {
      if (!panel) return;
      if (panel.imageId) ids.add(panel.imageId);
      if (panel.image) data.add(panel.image);
    });
  });
  return { ids, data };
};

// 削除してよいかの判定を返す。direct は常に削除、code 一致は未配置のものだけ削除する。
export const createImageDeletionFilter = ({ identity, sheetKeys = { ids: new Set(), data: new Set() } }) => (image) => {
  const kind = getImageDeletionMatchKind(image, identity);
  if (!kind) return null;
  if (kind === 'direct') return kind;
  const resolvedData = image.data || image.image || null;
  const isPlaced = (image.id && sheetKeys.ids.has(image.id)) || (resolvedData && sheetKeys.data.has(resolvedData));
  return isPlaced ? null : kind;
};

// 作業者フィルタ: workedBy 未記録の既存画像は全員に表示し、記録済みは本人のみに表示する。
export const isImageWorkedByUser = (image, userId) => {
  const workedBy = Array.isArray(image?.workedBy) ? image.workedBy : null;
  if (!workedBy || workedBy.length === 0) return true;
  if (!userId) return true;
  return workedBy.includes(userId);
};
