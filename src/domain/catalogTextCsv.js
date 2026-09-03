import {
  extractPdfCatalogDetails,
  extractPdfPriceFields,
  normalizePdfCropCode
} from './pdfCropImport.js';
import { buildBomCsvContent, escapeCsvCell } from '../lib/csv.js';

export const CATALOG_TEXT_CSV_HEADERS = Object.freeze([
  '画像ID',
  '画像名',
  '介援隊コード',
  '商品名',
  '商品名取得元',
  '品番',
  '品番候補',
  '品番抽出確度',
  'キャッチコピー',
  'キャッチコピー候補',
  '税込価格',
  '税抜価格',
  '価格候補',
  '価格抽出確度',
  '販売区分',
  '販売区分表示',
  '取扱記号',
  'デモ機',
  '仕様',
  '成分・原材料・栄養',
  '材質・素材',
  '元PDF',
  'カタログページ',
  'PDFページ',
  'コマサイズ',
  '元テキスト',
  'テキスト抽出バージョン',
  '文字省略',
  '登録日時'
]);

const LIST_SEPARATOR = '｜';

const joinList = (values) => (Array.isArray(values) ? values.filter(Boolean).join(LIST_SEPARATOR) : '');

const formatPriceCandidates = (candidates) => joinList((Array.isArray(candidates) ? candidates : []).map((candidate) => {
  const taxLabel = candidate?.taxType === 'excluding' ? '税抜' : '税込';
  return `${taxLabel}:${candidate?.amount ?? ''}${candidate?.text ? `(${candidate.text})` : ''}`;
}));

const formatCreatedAt = (createdAt) => {
  let date = null;
  if (typeof createdAt?.toDate === 'function') date = createdAt.toDate();
  else if (Number.isFinite(createdAt?.seconds)) date = new Date(createdAt.seconds * 1000);
  else if (typeof createdAt === 'string' || createdAt instanceof Date) date = new Date(createdAt);
  return date && !Number.isNaN(date.getTime()) ? date.toISOString() : '';
};

const formatAvailability = (value) => ({
  stock: '在庫',
  direct: '直送',
  mixed: '在庫・直送',
  unknown: '不明'
}[value] || '不明');

export const isCatalogTextExportImage = (image) => !!(
  image?.sourcePdfName
  || image?.sourceText
  || image?.catalogTextData
  || image?.textExtractionRect
);

const preferSavedList = (saved, analyzed, key) => (
  Array.isArray(saved?.[key]) && saved[key].length > 0 ? saved[key] : analyzed[key]
);

export const resolveCatalogTextExportData = (image = {}) => {
  const code = normalizePdfCropCode(image.catalogCode || image.code || image.name || '');
  const analyzedDetails = extractPdfCatalogDetails(image.sourceText || '', {
    code,
    productName: image.productName || ''
  });
  const savedDetails = image.catalogTextData || {};
  const details = {
    ...analyzedDetails,
    ...savedDetails,
    itemNumberCandidates: preferSavedList(savedDetails, analyzedDetails, 'itemNumberCandidates'),
    catchCopyCandidates: preferSavedList(savedDetails, analyzedDetails, 'catchCopyCandidates'),
    availabilityLabels: preferSavedList(savedDetails, analyzedDetails, 'availabilityLabels'),
    handlingMarkers: preferSavedList(savedDetails, analyzedDetails, 'handlingMarkers'),
    specifications: preferSavedList(savedDetails, analyzedDetails, 'specifications'),
    compositionDetails: preferSavedList(savedDetails, analyzedDetails, 'compositionDetails'),
    materialDetails: preferSavedList(savedDetails, analyzedDetails, 'materialDetails')
  };
  const analyzedPrices = extractPdfPriceFields(image.sourceText || '', code);
  const priceCandidates = Array.isArray(image.priceCandidates) && image.priceCandidates.length > 0
    ? image.priceCandidates
    : (analyzedPrices.priceCandidates || []);

  return {
    code,
    details,
    priceIncludingTax: Number.isFinite(image.priceIncludingTax)
      ? image.priceIncludingTax
      : (analyzedPrices.priceIncludingTax ?? ''),
    priceExcludingTax: Number.isFinite(image.priceExcludingTax)
      ? image.priceExcludingTax
      : (analyzedPrices.priceExcludingTax ?? ''),
    priceCandidates,
    priceExtractionConfidence: image.priceExtractionConfidence || analyzedPrices.priceExtractionConfidence || 'none'
  };
};

export const buildCatalogTextCsvRow = (image) => {
  const resolved = resolveCatalogTextExportData(image);
  const { details } = resolved;
  return [
    image.id || '',
    image.name || '',
    resolved.code,
    image.productName || '',
    image.productNameSource || '',
    details.itemNumber || '',
    joinList(details.itemNumberCandidates),
    details.itemNumberExtractionConfidence || '',
    details.catchCopy || '',
    joinList(details.catchCopyCandidates),
    resolved.priceIncludingTax,
    resolved.priceExcludingTax,
    formatPriceCandidates(resolved.priceCandidates),
    resolved.priceExtractionConfidence,
    formatAvailability(details.availability),
    joinList(details.availabilityLabels),
    joinList(details.handlingMarkers),
    details.hasDemoMarker ? 'あり' : 'なし',
    joinList(details.specifications),
    joinList(details.compositionDetails),
    joinList(details.materialDetails),
    image.sourcePdfName || '',
    image.sourcePage ?? '',
    image.pdfPageNumber ?? '',
    image.sizeType || '',
    image.sourceText || '',
    image.sourceTextVersion ?? '',
    image.sourceTextTruncated ? 'あり' : 'なし',
    formatCreatedAt(image.createdAt)
  ].map(escapeCsvCell).join(',');
};

export const buildCatalogTextCsvContent = (images = []) => {
  const exportImages = images.filter(isCatalogTextExportImage).sort((left, right) => {
    const leftPage = Number.isFinite(Number(left?.sourcePage)) ? Number(left.sourcePage) : Number.MAX_SAFE_INTEGER;
    const rightPage = Number.isFinite(Number(right?.sourcePage)) ? Number(right.sourcePage) : Number.MAX_SAFE_INTEGER;
    if (leftPage !== rightPage) return leftPage - rightPage;
    const leftCode = normalizePdfCropCode(left?.catalogCode || left?.code || left?.name || '');
    const rightCode = normalizePdfCropCode(right?.catalogCode || right?.code || right?.name || '');
    return leftCode.localeCompare(rightCode, 'ja', { numeric: true });
  });
  const rows = exportImages.map(buildCatalogTextCsvRow);
  return buildBomCsvContent(CATALOG_TEXT_CSV_HEADERS, rows, { escapeHeaders: true });
};

export const countCatalogTextExportImages = (images = []) => images.filter(isCatalogTextExportImage).length;
