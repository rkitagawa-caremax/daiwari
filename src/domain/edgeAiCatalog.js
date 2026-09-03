import { resolveCatalogTextExportData } from './catalogTextCsv.js';
import { normalizePdfCropCode } from './pdfCropImport.js';

export const EDGE_AI_MODEL_VERSION = 'ruri-v3-30m-int8-v1';

export const CATALOG_DIFF_FIELD_DEFINITIONS = Object.freeze([
  { key: 'name', label: '商品名', aliases: ['商品名', '品名', '名称', '商品名称'] },
  { key: 'itemNumber', label: '品番', aliases: ['品番', '型番', 'メーカー品番', '商品番号'] },
  { key: 'catchCopy', label: 'キャッチコピー', aliases: ['キャッチコピー', 'キャッチ', 'コピー', '商品説明', '説明文'] },
  { key: 'priceIncludingTax', label: '税込価格', aliases: ['税込価格', '価格(税込)', '税込', '税込み価格'] },
  { key: 'priceExcludingTax', label: '税抜価格', aliases: ['税抜価格', '価格(税抜)', '税別価格', '本体価格'] },
  { key: 'availability', label: '販売区分', aliases: ['販売区分', '在庫・直送', '在庫直送', '在庫区分'] },
  { key: 'lifecycleStatus', label: '販売状態', aliases: ['販売状態', '商品状態', '商品ステータス', '掲載状態', '廃盤・在庫限り', '廃盤在庫限り'] },
  { key: 'handlingMarkers', label: '取扱記号', aliases: ['取扱記号', '取扱い記号', '記号'] },
  { key: 'demoStatus', label: 'デモ機', aliases: ['デモ機', 'デモ機(D)', 'デモ', 'デモ機区分', 'デモ区分', 'デモ記号'] },
  { key: 'specifications', label: '仕様', aliases: ['仕様', '商品仕様', 'スペック'] },
  { key: 'compositionDetails', label: '成分・原材料・栄養', aliases: ['成分・原材料・栄養', '成分', '原材料', '栄養'] },
  { key: 'materialDetails', label: '材質・素材', aliases: ['材質・素材', '材質', '素材'] }
]);

const CODE_ALIASES = ['介援隊コード', '商品コード', 'コード', '介援隊CD', '介援隊ｺｰﾄﾞ'];
const LIST_SEPARATOR = ' / ';

const asArray = (value) => (
  Array.isArray(value) ? value.filter(Boolean).map((item) => String(item).trim()).filter(Boolean) : []
);

const displayValue = (value) => {
  if (Array.isArray(value)) return value.join(LIST_SEPARATOR);
  if (typeof value === 'boolean') return value ? 'あり' : 'なし';
  return String(value ?? '').trim();
};

const normalizeText = (value) => String(value ?? '')
  .normalize('NFKC')
  .toLowerCase()
  .replace(/[\s\u3000]+/g, '')
  .replace(/[、。,.・/／()（）「」『』【】{}:：;；'"!?！？￥¥]/g, '')
  .replaceAll('[', '')
  .replaceAll(']', '');

const normalizeHeader = (value) => normalizeText(value).replace(/[_-]/g, '');

const normalizePrice = (value) => {
  const text = String(value ?? '').normalize('NFKC').replace(/[,，\s￥¥円]/g, '');
  if (!text) return '';
  const match = text.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : String(value).trim();
};

const normalizeAvailability = (value) => {
  const normalized = normalizeText(value);
  if (!normalized) return '';
  const hasStock = /在庫|stock/.test(normalized);
  const hasDirect = /直送|direct/.test(normalized);
  if (hasStock && hasDirect) return 'mixed';
  if (hasStock) return 'stock';
  if (hasDirect) return 'direct';
  if (/受注生産/.test(normalized)) return 'made-to-order';
  return String(value).trim();
};

const normalizeLifecycleStatus = (value) => {
  const normalized = normalizeText(value);
  if (!normalized) return '';
  if (/廃盤|販売終了|終売/.test(normalized)) return '廃盤';
  if (/在庫限り|在庫かぎり|売切次第終了/.test(normalized)) return '在庫限り';
  if (/休止|一時停止/.test(normalized)) return '休止';
  if (/通常|継続|販売中/.test(normalized)) return '通常';
  return String(value).trim();
};

const normalizeDemoStatus = (value) => {
  const raw = String(value ?? '').normalize('NFKC').trim();
  if (!raw) return '';
  const compact = raw.replace(/[（）()\s]/g, '').toUpperCase();
  if (compact === 'D' || compact === 'デモ機' || compact === 'あり' || compact === '有' || compact === 'YES' || compact === 'TRUE' || compact === '1' || compact === '○') return 'D';
  if (compact === 'N') return 'N';
  if (compact === 'なし' || compact === '無' || compact === 'NO' || compact === 'FALSE' || compact === '0' || compact === '×' || compact === '-') return '';
  return raw;
};

const normalizeList = (value) => String(value ?? '')
  .split(/[|｜;；\n]|\s*[・●]\s*/)
  .map((item) => item.trim())
  .filter(Boolean);

const unique = (values) => [...new Set(values.filter(Boolean))];

const makeSearchText = (product) => unique([
  product.code,
  product.name,
  product.itemNumber,
  product.catchCopy,
  displayValue(product.availability),
  product.lifecycleStatus,
  ...product.handlingMarkers,
  product.demoStatus ? `デモ機 ${product.demoStatus}` : '',
  ...product.specifications,
  ...product.compositionDetails,
  ...product.materialDetails,
  product.sizeType,
  ...product.salesNames,
  ...product.salesSpecs,
  product.sourceText
]).join('。');

const resolveImageCode = (image) => normalizePdfCropCode(
  image?.catalogCode || image?.code || image?.name || ''
);

const buildAssignmentMaps = (sheets, genreLabels = {}) => {
  const byImageId = new Map();
  const byImageData = new Map();
  const byCode = new Map();

  const add = (map, key, assignment) => {
    if (!key) return;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(assignment);
  };

  (Array.isArray(sheets) ? sheets : []).forEach((sheet, sheetIndex) => {
    (Array.isArray(sheet?.panels) ? sheet.panels : []).forEach((panel, panelIndex) => {
      if (!panel || panel.hidden || (!panel.image && !panel.imageId)) return;
      const code = normalizePdfCropCode(panel.code || panel.originalName || '');
      const assignment = {
        sheetId: sheet.id,
        pageNumber: sheetIndex + 1,
        panelIndex,
        genre: genreLabels[sheet.genre] || sheet.genre || '未設定'
      };
      add(byImageId, panel.imageId, assignment);
      add(byImageData, panel.image, assignment);
      add(byCode, code, assignment);
    });
  });

  return { byImageId, byImageData, byCode };
};

const getImageAssignments = (image, code, maps) => {
  const values = [
    ...(maps.byImageId.get(image?.id) || []),
    ...(maps.byImageData.get(image?.data) || []),
    ...(maps.byCode.get(code) || [])
  ];
  const seen = new Set();
  return values.filter((entry) => {
    const key = `${entry.sheetId}:${entry.panelIndex}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

export const buildEdgeCatalogProducts = ({ images = [], sheets = [], salesData = {}, genres = [] } = {}) => {
  const genreLabels = Object.fromEntries((genres || []).map((genre) => [genre.id, genre.label]));
  const assignmentMaps = buildAssignmentMaps(sheets, genreLabels);
  const products = [];

  (Array.isArray(images) ? images : []).forEach((image, imageIndex) => {
    if (!image) return;
    const resolved = resolveCatalogTextExportData(image);
    const code = resolved.code || resolveImageCode(image);
    if (!code && !image.productName && !image.sourceText) return;
    const details = resolved.details || {};
    const salesRows = code && Array.isArray(salesData?.[code]) ? salesData[code] : [];
    const assignments = getImageAssignments(image, code, assignmentMaps);
    const product = {
      id: image.id || `${code || 'unknown'}-${imageIndex}`,
      imageId: image.id || null,
      imageData: image.data || null,
      code,
      name: image.productName || salesRows.find((row) => row?.name)?.name || '',
      itemNumber: details.itemNumber || '',
      catchCopy: details.catchCopy || '',
      priceIncludingTax: resolved.priceIncludingTax ?? '',
      priceExcludingTax: resolved.priceExcludingTax ?? '',
      priceExtractionConfidence: resolved.priceExtractionConfidence || 'none',
      availability: details.availability || '',
      lifecycleStatus: details.lifecycleStatus || '',
      handlingMarkers: asArray(details.handlingMarkers),
      hasDemoMarker: details.hasDemoMarker === true,
      demoStatus: asArray(details.handlingMarkers).map((marker) => marker.match(/^\(([DN])\)$/)?.[1]).find(Boolean)
        || (details.hasDemoMarker === true ? 'D' : ''),
      specifications: asArray(details.specifications),
      compositionDetails: asArray(details.compositionDetails),
      materialDetails: asArray(details.materialDetails),
      sizeType: image.sizeType || '',
      sourceText: String(image.sourceText || '').trim(),
      sourcePdfName: image.sourcePdfName || '',
      sourcePage: image.sourcePage ?? '',
      assignments,
      salesCount: salesRows.reduce((total, row) => total + (Number(row?.count) || 0), 0),
      salesNames: unique(salesRows.map((row) => String(row?.name || '').trim())),
      salesSpecs: unique(salesRows.map((row) => String(row?.spec || '').trim()))
    };
    product.searchText = makeSearchText(product);
    products.push(product);
  });

  return products;
};

export const createEdgeCatalogFingerprint = (products = []) => {
  let hash = 2166136261;
  [...products]
    .sort((left, right) => String(left.id).localeCompare(String(right.id)))
    .forEach((product) => {
      const value = `${product.id}\u0000${product.searchText}`;
      for (let index = 0; index < value.length; index += 1) {
        hash ^= value.charCodeAt(index);
        hash = Math.imul(hash, 16777619);
      }
    });
  return `${EDGE_AI_MODEL_VERSION}:${products.length}:${(hash >>> 0).toString(36)}`;
};

const ngrams = (value, size = 2) => {
  const normalized = normalizeText(value);
  if (!normalized) return new Set();
  if (normalized.length <= size) return new Set([normalized]);
  const result = new Set();
  for (let index = 0; index <= normalized.length - size; index += 1) {
    result.add(normalized.slice(index, index + size));
  }
  return result;
};

const setSimilarity = (left, right) => {
  if (!left.size || !right.size) return 0;
  let overlap = 0;
  left.forEach((value) => { if (right.has(value)) overlap += 1; });
  return overlap / Math.max(1, left.size + right.size - overlap);
};

export const cosineSimilarity = (left = [], right = []) => {
  if (!left.length || left.length !== right.length) return 0;
  let dot = 0;
  let leftNorm = 0;
  let rightNorm = 0;
  for (let index = 0; index < left.length; index += 1) {
    dot += left[index] * right[index];
    leftNorm += left[index] ** 2;
    rightNorm += right[index] ** 2;
  }
  return leftNorm && rightNorm ? dot / Math.sqrt(leftNorm * rightNorm) : 0;
};

const lexicalScore = (product, query) => {
  const normalizedQuery = normalizeText(query);
  if (!normalizedQuery) return 0;
  const normalizedCode = normalizeText(product.code);
  const normalizedName = normalizeText(product.name);
  const normalizedDocument = normalizeText(product.searchText);
  const exactCode = normalizedCode && normalizedCode === normalizedQuery ? 1 : 0;
  const prefixCode = normalizedCode && normalizedCode.startsWith(normalizedQuery) ? 0.82 : 0;
  const nameMatch = normalizedName.includes(normalizedQuery) ? 0.76 : 0;
  const documentMatch = normalizedDocument.includes(normalizedQuery) ? 0.62 : 0;
  const fuzzy = setSimilarity(ngrams(normalizedQuery), ngrams(product.searchText));
  return Math.max(exactCode, prefixCode, nameMatch, documentMatch, fuzzy * 0.72);
};

export const rankEdgeCatalogProducts = (
  products = [],
  query = '',
  { queryVector = null, vectorsById = null, limit = 30 } = {}
) => products
  .map((product) => {
    const lexical = lexicalScore(product, query);
    const vector = vectorsById?.[product.id];
    const semantic = queryVector && vector ? Math.max(0, cosineSimilarity(queryVector, vector)) : 0;
    const score = queryVector ? Math.max(lexical, semantic * 0.86 + lexical * 0.14) : lexical;
    return { product, score, lexicalScore: lexical, semanticScore: semantic };
  })
  .filter((result) => result.score > 0.04)
  .sort((left, right) => right.score - left.score || right.product.salesCount - left.product.salesCount)
  .slice(0, limit);

const sharedValueScore = (left, right) => {
  const listFields = ['handlingMarkers', 'specifications', 'compositionDetails', 'materialDetails'];
  let score = 0;
  listFields.forEach((field) => {
    score += setSimilarity(ngrams(displayValue(left[field])), ngrams(displayValue(right[field]))) * 0.04;
  });
  if (left.sizeType && left.sizeType === right.sizeType) score += 0.05;
  if (left.availability && left.availability === right.availability) score += 0.03;
  return Math.min(score, 0.18);
};

export const rankSimilarEdgeCatalogProducts = (
  products = [],
  sourceProduct,
  { vectorsById = null, limit = 12 } = {}
) => {
  if (!sourceProduct) return [];
  const sourceVector = vectorsById?.[sourceProduct.id];
  return products
    .filter((candidate) => candidate.id !== sourceProduct.id)
    .map((candidate) => {
      const semantic = sourceVector && vectorsById?.[candidate.id]
        ? Math.max(0, cosineSimilarity(sourceVector, vectorsById[candidate.id]))
        : setSimilarity(ngrams(sourceProduct.searchText), ngrams(candidate.searchText));
      const score = Math.min(1, semantic * 0.9 + sharedValueScore(sourceProduct, candidate));
      return { product: candidate, score, semanticScore: semantic };
    })
    .filter((result) => result.score > 0.08)
    .sort((left, right) => right.score - left.score || right.product.salesCount - left.product.salesCount)
    .slice(0, limit);
};

export const parseCsvRecords = (csvText = '') => {
  const text = String(csvText).replace(/^\uFEFF/, '');
  const records = [];
  let row = [];
  let cell = '';
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === '"') {
      if (quoted && text[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (character === ',' && !quoted) {
      row.push(cell.trim());
      cell = '';
    } else if ((character === '\n' || character === '\r') && !quoted) {
      if (character === '\r' && text[index + 1] === '\n') index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) records.push(row);
      row = [];
      cell = '';
    } else {
      cell += character;
    }
  }
  row.push(cell.trim());
  if (row.some(Boolean)) records.push(row);
  return records;
};

const findHeaderIndex = (headers, aliases) => {
  const normalizedAliases = aliases.map(normalizeHeader);
  return headers.findIndex((header) => normalizedAliases.includes(normalizeHeader(header)));
};

const findCatalogHeaderRowIndex = (records = []) => {
  let best = { index: -1, score: -1 };
  records.slice(0, 30).forEach((record, index) => {
    const headers = Array.isArray(record) ? record : [];
    if (findHeaderIndex(headers, CODE_ALIASES) < 0) return;
    const recognizedCount = CATALOG_DIFF_FIELD_DEFINITIONS.reduce((count, field) => (
      count + (findHeaderIndex(headers, field.aliases) >= 0 ? 1 : 0)
    ), 0);
    const score = recognizedCount * 100 + headers.filter((value) => String(value ?? '').trim()).length;
    if (score > best.score) best = { index, score };
  });
  return best.index;
};

export const parseCatalogSnapshotRecords = (inputRecords = []) => {
  const records = (Array.isArray(inputRecords) ? inputRecords : []).map((record) => (
    (Array.isArray(record) ? record : []).map((value) => (
      value instanceof Date ? value.toISOString() : value ?? ''
    ))
  ));
  if (!records.length) throw new Error('ファイルにデータがありません。');
  const headerRowIndex = findCatalogHeaderRowIndex(records);
  if (headerRowIndex < 0) throw new Error('「介援隊コード」または「商品コード」列が見つかりません。');
  const headers = records[headerRowIndex];
  const codeIndex = findHeaderIndex(headers, CODE_ALIASES);
  const fieldIndexes = Object.fromEntries(CATALOG_DIFF_FIELD_DEFINITIONS.map((field) => [
    field.key,
    findHeaderIndex(headers, field.aliases)
  ]));
  const byCode = new Map();
  let unreadableRows = 0;
  let duplicateCodes = 0;

  records.slice(headerRowIndex + 1).forEach((record) => {
    const code = normalizePdfCropCode(record[codeIndex] || '');
    if (!code) {
      if (record.some(Boolean)) unreadableRows += 1;
      return;
    }
    if (byCode.has(code)) duplicateCodes += 1;
    const item = { code };
    CATALOG_DIFF_FIELD_DEFINITIONS.forEach((field) => {
      const index = fieldIndexes[field.key];
      let value = index >= 0 ? record[index] ?? '' : '';
      if (field.key.startsWith('price')) value = normalizePrice(value);
      else if (field.key === 'availability') value = normalizeAvailability(value);
      else if (field.key === 'lifecycleStatus') value = normalizeLifecycleStatus(value);
      else if (field.key === 'demoStatus') value = normalizeDemoStatus(value);
      else if (['handlingMarkers', 'specifications', 'compositionDetails', 'materialDetails'].includes(field.key)) {
        value = normalizeList(value);
      } else value = String(value).trim();
      item[field.key] = value;
    });
    item.hasDemoMarker = item.demoStatus === 'D';
    byCode.set(code, item);
  });

  return {
    headers,
    items: [...byCode.values()],
    duplicateCodes,
    unreadableRows,
    headerRowIndex,
    recognizedFields: CATALOG_DIFF_FIELD_DEFINITIONS.filter((field) => fieldIndexes[field.key] >= 0).map((field) => field.key)
  };
};

export const parseCatalogSnapshotCsv = (csvText = '') => parseCatalogSnapshotRecords(parseCsvRecords(csvText));

const getChangeSeverity = (fieldKey) => {
  if (['priceIncludingTax', 'priceExcludingTax', 'availability', 'lifecycleStatus', 'handlingMarkers', 'demoStatus', 'hasDemoMarker'].includes(fieldKey)) return 'high';
  if (['itemNumber', 'specifications', 'compositionDetails', 'materialDetails'].includes(fieldKey)) return 'medium';
  return 'normal';
};

const displayFieldValue = (fieldKey, value) => {
  if (fieldKey === 'availability') {
    return ({ stock: '在庫', direct: '直送', mixed: '在庫・直送', 'made-to-order': '受注生産' }[value]) || displayValue(value);
  }
  if (fieldKey === 'demoStatus') return value ? `(${String(value).replace(/[（）()]/g, '')})` : 'なし';
  return displayValue(value);
};

const areFieldValuesEqual = (fieldKey, left, right) => (
  normalizeText(displayFieldValue(fieldKey, left)) === normalizeText(displayFieldValue(fieldKey, right))
);

export const compareCatalogSnapshots = (oldItems = [], newItems = [], options = {}) => {
  const oldByCode = new Map(oldItems.map((item) => [item.code, item]));
  const newByCode = new Map(newItems.map((item) => [item.code, item]));
  const missingMeansRemoved = options.missingMeansRemoved !== false;
  const requestedFields = Array.isArray(options.fields) ? new Set(options.fields) : null;
  const allCodes = [...new Set([
    ...(missingMeansRemoved ? oldByCode.keys() : []),
    ...newByCode.keys()
  ])].sort((a, b) => a.localeCompare(b, 'ja', { numeric: true }));

  return allCodes.map((code) => {
    const before = oldByCode.get(code);
    const after = newByCode.get(code);
    if (!before) return { code, status: 'added', before: null, after, changes: [] };
    if (!after) return { code, status: 'removed', before, after: null, changes: [] };
    const changes = CATALOG_DIFF_FIELD_DEFINITIONS
      .filter((field) => !requestedFields || requestedFields.has(field.key))
      .filter((field) => !areFieldValuesEqual(field.key, before[field.key], after[field.key]))
      .map((field) => ({
        key: field.key,
        label: field.label,
        before: displayFieldValue(field.key, before[field.key]),
        after: displayFieldValue(field.key, after[field.key]),
        severity: getChangeSeverity(field.key),
        confidence: field.key.startsWith('price') ? (before.priceExtractionConfidence || 'unknown') : 'high'
      }));
    return { code, status: changes.length ? 'modified' : 'unchanged', before, after, changes };
  });
};

export const buildCatalogChangeSet = ({ products = [], snapshot = {}, fileName = '', sheetName = '' } = {}) => {
  const diffs = compareCatalogSnapshots(products, snapshot.items || [], {
    fields: snapshot.recognizedFields || [],
    missingMeansRemoved: false
  });
  const changedDiffs = diffs.filter((diff) => diff.status !== 'unchanged');
  const displayableDiffs = changedDiffs.filter((diff) => diff.status === 'modified' && diff.changes.length > 0);
  return {
    version: 1,
    fileName,
    sheetName,
    importedAt: new Date().toISOString(),
    recognizedFields: [...(snapshot.recognizedFields || [])],
    unreadableRows: snapshot.unreadableRows || 0,
    duplicateCodes: snapshot.duplicateCodes || 0,
    summary: summarizeCatalogDiff(diffs),
    diffs: changedDiffs,
    displayCount: displayableDiffs.length,
    byCode: Object.fromEntries(displayableDiffs.map((diff) => [diff.code, diff]))
  };
};

export const summarizeCatalogDiff = (diffs = []) => diffs.reduce((summary, diff) => {
  summary[diff.status] += 1;
  return summary;
}, { added: 0, removed: 0, modified: 0, unchanged: 0 });
