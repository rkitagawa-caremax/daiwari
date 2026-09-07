import { parseCSVLine } from '../lib/csv.js';
import { normalizeCode } from './productCodes.js';

export const SALES_DATA_HEADER_ROW_COUNT = 2;
export const SALES_DATA_CHUNK_SIZE = 1000;
export const SALES_DATA_CHUNK_MAX_BYTES = 640 * 1024;
export const SALES_DATA_WRITE_BATCH_SIZE = 5;

const normalizeSalesHeader = (value) => String(value ?? '')
  .normalize('NFKC')
  .replace(/[\s\u3000]/g, '');

const parseSalesValue = (value) => {
  const normalized = String(value ?? '')
    .normalize('NFKC')
    .replace(/[,，\s]/g, '');
  if (!normalized) return 0;
  const negativeMatch = normalized.match(/^\(([-+]?\d+(?:\.\d+)?)\)$/);
  const parsed = Number(negativeMatch ? `-${negativeMatch[1]}` : normalized);
  return Number.isFinite(parsed) ? parsed : (parseInt(normalized, 10) || 0);
};

const SALES_COLUMN_ALIASES = Object.freeze({
  code: ['介援隊コード', '介援隊CD', '商品コード', '品目コード', 'コード'],
  name: ['商品名', '品名', '商品名称', '名称'],
  spec: ['規格', '仕様', '商品規格'],
  quantity: ['販売数量', '売上数量', '売上数', '販売数', '数量合計', '出荷数量'],
  salesAmount: ['売上額', '売上金額', '販売金額', '売上高', '税抜売上額'],
  grossProfitAmount: ['粗利額', '粗利益額', '粗利益', '荒利額', '荒利益額']
});

const resolveHeaderColumn = (columnHeaders, aliases, fallbackIndex = -1) => {
  const normalizedAliases = aliases.map(normalizeSalesHeader);
  const candidates = columnHeaders
    .map((header, columnIndex) => {
      const normalized = normalizeSalesHeader(header);
      const matchedAlias = normalizedAliases.find((alias) => normalized === alias || normalized.includes(alias));
      if (!matchedAlias) return null;
      const isMonthly = /(?:^|[^0-9])(1[0-2]|0?[1-9])月/.test(normalized)
        || /(?:19|20)\d{2}[-/.](1[0-2]|0?[1-9])/.test(normalized);
      const score = (normalized === matchedAlias ? 100 : 60)
        + (/合計|累計|総計|年度計|期間計/.test(normalized) ? 30 : 0)
        - (isMonthly ? 50 : 0);
      return { columnIndex, score };
    })
    .filter(Boolean)
    .sort((left, right) => right.score - left.score || right.columnIndex - left.columnIndex);
  return candidates[0]?.columnIndex ?? fallbackIndex;
};

export const resolveSalesColumnSchema = (headerRows = []) => {
  const rows = (Array.isArray(headerRows) ? headerRows : []).map((row) => (
    Array.isArray(row) ? row : []
  ));
  const columnCount = Math.max(0, ...rows.map((row) => row.length));
  const columnHeaders = Array.from({ length: columnCount }, (_, columnIndex) => (
    rows.map((row) => row[columnIndex]).filter(Boolean).join(' ')
  ));
  const quantityIndex = resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.quantity);
  const salesAmountIndex = resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.salesAmount);
  const grossProfitAmountIndex = resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.grossProfitAmount);
  const hasRecognizedMetric = quantityIndex >= 0 || salesAmountIndex >= 0 || grossProfitAmountIndex >= 0;
  return {
    codeIndex: resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.code, 3),
    nameIndex: resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.name, 1),
    specIndex: resolveHeaderColumn(columnHeaders, SALES_COLUMN_ALIASES.spec, 2),
    // 列名を認識できない旧CSVだけ、従来仕様の18列目を数量として扱う。
    quantityIndex: hasRecognizedMetric ? quantityIndex : 17,
    salesAmountIndex,
    grossProfitAmountIndex
  };
};

const isCodeHeaderCell = (value) => {
  const normalized = normalizeSalesHeader(value);
  return SALES_COLUMN_ALIASES.code.some((alias) => {
    const normalizedAlias = normalizeSalesHeader(alias);
    return normalized === normalizedAlias || normalized.includes(normalizedAlias);
  });
};

export const resolveSalesCsvLayout = (rows = [], fallbackHeaderRowCount = SALES_DATA_HEADER_ROW_COUNT) => {
  const normalizedRows = Array.isArray(rows) ? rows : [];
  const probeCount = Math.min(normalizedRows.length, Math.max(8, fallbackHeaderRowCount));
  const parsedProbeRows = normalizedRows.slice(0, probeCount).map((row) => parseCSVLine(row));
  const detectedHeaderIndex = parsedProbeRows.findIndex((row) => row.some(isCodeHeaderCell));
  if (detectedHeaderIndex < 0) {
    return {
      dataStartIndex: fallbackHeaderRowCount,
      headerRows: parsedProbeRows.slice(0, fallbackHeaderRowCount),
      detectedHeaderIndex: null
    };
  }

  let dataStartIndex = detectedHeaderIndex + 1;
  let headerRows = parsedProbeRows.slice(detectedHeaderIndex, dataStartIndex);
  let schema = resolveSalesColumnSchema(headerRows);
  // 年度行＋月行など、商品コードが空の連続行は複数段ヘッダーとして取り込む。
  while (dataStartIndex < parsedProbeRows.length && dataStartIndex <= detectedHeaderIndex + 3) {
    const candidate = parsedProbeRows[dataStartIndex];
    if (String(candidate?.[schema.codeIndex] || '').trim()) break;
    headerRows = [...headerRows, candidate];
    dataStartIndex += 1;
    schema = resolveSalesColumnSchema(headerRows);
  }

  return { dataStartIndex, headerRows, detectedHeaderIndex };
};

export const resolveSalesMonthColumns = (headerRows = []) => {
  const rows = (Array.isArray(headerRows) ? headerRows : []).map((row) => (
    Array.isArray(row) ? row : []
  ));
  const columnCount = Math.max(0, ...rows.map((row) => row.length));
  const columns = [];
  let carriedYear = null;
  let previousMonth = null;
  let carriedMetric = null;

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    // 現行帳票の0〜3列は商品情報。帳票タイトル内の年月を売上列と誤認しない。
    if (columnIndex <= 3) continue;
    const headerParts = rows
      .map((row) => normalizeSalesHeader(row[columnIndex]))
      .filter(Boolean);
    const header = headerParts.join(' ');
    const firstHeader = normalizeSalesHeader(rows[0]?.[columnIndex]);
    if (SALES_COLUMN_ALIASES.grossProfitAmount.some((alias) => firstHeader.includes(normalizeSalesHeader(alias)))) carriedMetric = 'grossProfitAmount';
    else if (SALES_COLUMN_ALIASES.salesAmount.some((alias) => firstHeader.includes(normalizeSalesHeader(alias)))) carriedMetric = 'salesAmount';
    else if (SALES_COLUMN_ALIASES.quantity.some((alias) => firstHeader.includes(normalizeSalesHeader(alias)))) carriedMetric = 'quantity';
    const explicitYear = header.match(/(?:19|20)\d{2}/)?.[0];
    if (explicitYear) carriedYear = Number(explicitYear);
    if (/合計|累計|平均|前年差|前年比/.test(header)) continue;
    if (carriedMetric === 'salesAmount' || carriedMetric === 'grossProfitAmount') continue;

    const slashMonthMatch = header.match(/(?:19|20)\d{2}[-/.](1[0-2]|0?[1-9])(?:月)?/);
    const japaneseMonthMatch = header.match(/(?:^|[^0-9])(1[0-2]|0?[1-9])月/);
    const month = Number(slashMonthMatch?.[1] || japaneseMonthMatch?.[1] || 0);
    if (month < 1 || month > 12) continue;

    if (!explicitYear && carriedYear && previousMonth && month < previousMonth) carriedYear += 1;
    const year = explicitYear ? Number(explicitYear) : carriedYear;
    const paddedMonth = String(month).padStart(2, '0');
    columns.push({
      columnIndex,
      key: year ? `${year}-${paddedMonth}` : `month-${paddedMonth}-${columnIndex}`,
      label: year ? `${String(year).slice(-2)}/${month}` : `${month}月`,
      year: year || null,
      month
    });
    previousMonth = month;
  }

  return columns;
};

export const buildMonthlySalesSeries = (items = []) => {
  const totals = new Map();
  const normalizedItems = Array.isArray(items) ? items : [];
  const sharedLabels = normalizedItems.find((item) => Array.isArray(item?.monthlyLabels))?.monthlyLabels || [];
  normalizedItems.forEach((item) => {
    (Array.isArray(item?.monthlySales) ? item.monthlySales : []).forEach((entry, index) => {
      const isCompactEntry = typeof entry === 'number';
      const label = isCompactEntry ? (item.monthlyLabels?.[index] || sharedLabels[index] || `${index + 1}月`) : (entry?.label || '');
      const key = isCompactEntry ? `compact-month-${index}` : (entry?.key || `${entry?.label || '月'}-${index}`);
      if (!totals.has(key)) {
        totals.set(key, {
          key,
          label,
          year: isCompactEntry ? null : (entry?.year ?? null),
          month: isCompactEntry ? null : (entry?.month ?? null),
          count: 0
        });
      }
      totals.get(key).count += Number(isCompactEntry ? entry : entry?.count) || 0;
    });
  });
  return [...totals.values()];
};

export const summarizeGrossProfitRows = (items = []) => {
  const rows = Array.isArray(items) ? items : [];
  const hasSalesAmount = rows.some((item) => Object.hasOwn(item || {}, 'salesAmount'));
  const hasGrossProfitAmount = rows.some((item) => Object.hasOwn(item || {}, 'grossProfitAmount'));
  const salesAmount = rows.reduce((sum, item) => sum + (Number(item?.salesAmount) || 0), 0);
  const grossProfitAmount = rows.reduce((sum, item) => sum + (Number(item?.grossProfitAmount) || 0), 0);
  const grossMargin = hasSalesAmount && hasGrossProfitAmount && salesAmount !== 0
    ? grossProfitAmount / salesAmount
    : null;
  return {
    hasSalesAmount,
    hasGrossProfitAmount,
    salesAmount,
    grossProfitAmount,
    grossMargin,
    chartRatio: grossMargin == null ? 0 : Math.max(0, Math.min(1, grossMargin))
  };
};

// 売上CSVのUIや保存先に依存しない解析処理。
export const parseSalesCsvContent = (
  csvText,
  { headerRowCount = SALES_DATA_HEADER_ROW_COUNT, metricType = 'auto' } = {}
) => {
  const normalizedText = String(csvText ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rows = normalizedText.split('\n');
  const layout = resolveSalesCsvLayout(rows, headerRowCount);
  const parsedHeaderRows = layout.headerRows;
  const columnSchema = resolveSalesColumnSchema(parsedHeaderRows);
  const validMetricType = ['quantity', 'salesAmount', 'grossProfitAmount'].includes(metricType)
    ? metricType
    : 'auto';
  const explicitMetricIndex = validMetricType === 'auto'
    ? -1
    : (columnSchema[`${validMetricType}Index`] >= 0 ? columnSchema[`${validMetricType}Index`] : 17);
  const monthColumns = validMetricType === 'auto' || validMetricType === 'quantity'
    ? resolveSalesMonthColumns(parsedHeaderRows)
    : [];
  const salesData = {};

  rows.slice(layout.dataStartIndex).forEach((row) => {
    if (!row.trim()) return;
    const columns = parseCSVLine(row);
    if (columns.length <= columnSchema.codeIndex) return;

    const rawCode = columns[columnSchema.codeIndex];
    if (!rawCode) return;

    const code = normalizeCode(rawCode);
    const monthlySales = monthColumns.map((monthColumn) => parseSalesValue(columns[monthColumn.columnIndex]));
    if (!salesData[code]) salesData[code] = [];
    const item = {
      name: columns[columnSchema.nameIndex] || '',
      spec: columns[columnSchema.specIndex] || ''
    };
    if (validMetricType === 'auto') {
      if (columnSchema.quantityIndex >= 0 && columnSchema.quantityIndex < columns.length) {
        item.count = parseSalesValue(columns[columnSchema.quantityIndex]);
      } else if (monthlySales.length > 0) {
        item.count = monthlySales.reduce((sum, value) => sum + value, 0);
      }
      if (columnSchema.salesAmountIndex >= 0 && columnSchema.salesAmountIndex < columns.length) {
        item.salesAmount = parseSalesValue(columns[columnSchema.salesAmountIndex]);
      }
      if (columnSchema.grossProfitAmountIndex >= 0 && columnSchema.grossProfitAmountIndex < columns.length) {
        item.grossProfitAmount = parseSalesValue(columns[columnSchema.grossProfitAmountIndex]);
      }
    } else if (explicitMetricIndex >= 0 && explicitMetricIndex < columns.length) {
      item[validMetricType === 'quantity' ? 'count' : validMetricType] = parseSalesValue(columns[explicitMetricIndex]);
    } else if (validMetricType === 'quantity' && monthlySales.length > 0) {
      item.count = monthlySales.reduce((sum, value) => sum + value, 0);
    }
    if ((validMetricType === 'auto' || validMetricType === 'quantity') && monthlySales.length > 0) {
      item.monthlySales = monthlySales;
      // 月ラベルはコードごとに1度だけ保存し、Firestoreの通信量を抑える。
      if (salesData[code].length === 0) item.monthlyLabels = monthColumns.map((monthColumn) => monthColumn.label);
    }
    salesData[code].push(item);
  });

  return salesData;
};

const metricFields = Object.freeze({
  quantity: ['count', 'monthlySales', 'monthlyLabels'],
  salesAmount: ['salesAmount'],
  grossProfitAmount: ['grossProfitAmount']
});

const normalizeRowIdentity = (row) => [row?.name, row?.spec]
  .map((value) => String(value || '').normalize('NFKC').replace(/[\s\u3000]/g, '').toLowerCase())
  .join('|');

// 別ファイルで届く3指標を商品コード・名称・規格で突合し、選択した指標だけ更新する。
export const mergeSalesMetricData = (existingData = {}, importedData = {}, metricType = 'quantity') => {
  const fields = metricFields[metricType] || metricFields.quantity;
  const merged = Object.fromEntries(Object.entries(existingData || {}).map(([code, rows]) => [
    code,
    (Array.isArray(rows) ? rows : []).map((row) => {
      const next = { ...row };
      fields.forEach((field) => delete next[field]);
      return next;
    })
  ]));

  Object.entries(importedData || {}).forEach(([code, importedRows]) => {
    const currentRows = merged[code] || [];
    const identityIndexes = new Map();
    currentRows.forEach((row, index) => {
      const key = normalizeRowIdentity(row);
      if (key !== '|' && !identityIndexes.has(key)) identityIndexes.set(key, index);
    });
    (Array.isArray(importedRows) ? importedRows : []).forEach((incoming, incomingIndex) => {
      const identity = normalizeRowIdentity(incoming);
      const matchedIndex = identity !== '|' && identityIndexes.has(identity)
        ? identityIndexes.get(identity)
        : (currentRows[incomingIndex] ? incomingIndex : -1);
      const next = matchedIndex >= 0
        ? { ...currentRows[matchedIndex] }
        : { name: incoming?.name || '', spec: incoming?.spec || '' };
      if (!next.name && incoming?.name) next.name = incoming.name;
      if (!next.spec && incoming?.spec) next.spec = incoming.spec;
      fields.forEach((field) => {
        if (Object.hasOwn(incoming || {}, field)) next[field] = incoming[field];
        else delete next[field];
      });
      if (matchedIndex >= 0) currentRows[matchedIndex] = next;
      else {
        currentRows.push(next);
        if (identity !== '|') identityIndexes.set(identity, currentRows.length - 1);
      }
    });
    merged[code] = currentRows;
  });

  Object.keys(merged).forEach((code) => {
    merged[code] = merged[code].filter((row) => ['count', 'salesAmount', 'grossProfitAmount'].some((field) => Object.hasOwn(row || {}, field)));
    if (merged[code].length === 0) delete merged[code];
  });

  return merged;
};

export const summarizeSalesDataMetrics = (salesData = {}) => {
  const summary = {
    quantity: { codes: 0, total: 0 },
    salesAmount: { codes: 0, total: 0 },
    grossProfitAmount: { codes: 0, total: 0 }
  };
  Object.values(salesData || {}).forEach((items) => {
    const rows = Array.isArray(items) ? items : [];
    [
      ['quantity', 'count'],
      ['salesAmount', 'salesAmount'],
      ['grossProfitAmount', 'grossProfitAmount']
    ].forEach(([summaryKey, valueKey]) => {
      const matchedRows = rows.filter((item) => Object.hasOwn(item || {}, valueKey));
      if (matchedRows.length === 0) return;
      summary[summaryKey].codes += 1;
      summary[summaryKey].total += matchedRows.reduce((sum, item) => sum + (Number(item?.[valueKey]) || 0), 0);
    });
  });
  return summary;
};

export const validateSalesMetricImport = (importedData = {}, existingData = {}, metricType = 'quantity') => {
  const valueKey = metricType === 'salesAmount'
    ? 'salesAmount'
    : metricType === 'grossProfitAmount' ? 'grossProfitAmount' : 'count';
  const metricLabel = metricType === 'salesAmount'
    ? '売上額'
    : metricType === 'grossProfitAmount' ? '粗利額' : '販売数量';
  let importedRows = 0;
  let importedCodes = 0;
  let nonZeroRows = 0;
  Object.values(importedData || {}).forEach((rows) => {
    if (!Array.isArray(rows)) return;
    let codeHasMetric = false;
    rows.forEach((row) => {
      if (!Object.hasOwn(row || {}, valueKey)) return;
      codeHasMetric = true;
      importedRows += 1;
      if (Number(row?.[valueKey]) !== 0) nonZeroRows += 1;
    });
    if (codeHasMetric) importedCodes += 1;
  });
  if (importedCodes === 0 || importedRows === 0) {
    throw new Error(`${metricLabel}列または商品コードを認識できませんでした。CSVのヘッダーと列位置を確認してください。既存データは変更されていません。`);
  }
  if (nonZeroRows === 0) {
    throw new Error(`${metricLabel}がすべて空欄または0です。誤った列の取り込みを防ぐため保存を中止しました。既存データは変更されていません。`);
  }

  let existingCodes = 0;
  Object.values(existingData || {}).forEach((rows) => {
    if (Array.isArray(rows) && rows.some((row) => Object.hasOwn(row || {}, valueKey))) existingCodes += 1;
  });
  if (existingCodes >= 20 && importedCodes < existingCodes * 0.5) {
    throw new Error(`${metricLabel}の認識商品数が既存データの半数未満です（${importedCodes.toLocaleString()} / ${existingCodes.toLocaleString()}商品）。期間やCSV形式を確認してください。既存データは変更されていません。`);
  }
  return { metricType, valueKey, importedCodes, importedRows, nonZeroRows, existingCodes };
};

export const splitSalesDataIntoChunks = (
  salesData,
  chunkSize = SALES_DATA_CHUNK_SIZE,
  maxBytes = SALES_DATA_CHUNK_MAX_BYTES
) => {
  const entries = Object.entries(salesData || {});
  const chunks = [];
  const encoder = new TextEncoder();
  let pendingEntries = [];
  let pendingBytes = 2;

  const flush = () => {
    if (pendingEntries.length === 0) return;
    chunks.push(Object.fromEntries(pendingEntries));
    pendingEntries = [];
    pendingBytes = 2;
  };

  entries.forEach(([code, items]) => {
    const entryBytes = encoder.encode(JSON.stringify({ [code]: items })).byteLength;
    if (entryBytes > maxBytes) {
      throw new Error(`介援隊コード ${code} の実績明細が大きすぎるため保存できません。`);
    }
    if (pendingEntries.length > 0 && (pendingEntries.length >= chunkSize || pendingBytes + entryBytes > maxBytes)) flush();
    pendingEntries.push([code, items]);
    pendingBytes += entryBytes;
  });
  flush();
  return chunks;
};

export const mergeSerializedSalesChunks = (serializedChunks, { onParseError } = {}) => {
  const salesData = {};
  (serializedChunks || []).forEach((serializedChunk, index) => {
    if (!serializedChunk) return;
    try {
      Object.assign(salesData, JSON.parse(serializedChunk));
    } catch (error) {
      onParseError?.(error, index);
    }
  });
  return salesData;
};

const readSalesChunkDocumentData = (snapshotDoc) => {
  if (typeof snapshotDoc?.data === 'function') return snapshotDoc.data() || {};
  return snapshotDoc?.data || {};
};

const sortSalesChunkDocuments = (documents) => [...documents].sort((left, right) => {
  const leftData = readSalesChunkDocumentData(left);
  const rightData = readSalesChunkDocumentData(right);
  const leftIndex = Number(leftData.chunkIndex);
  const rightIndex = Number(rightData.chunkIndex);
  if (Number.isFinite(leftIndex) && Number.isFinite(rightIndex) && leftIndex !== rightIndex) {
    return leftIndex - rightIndex;
  }
  return String(left?.id || '').localeCompare(String(right?.id || ''), undefined, { numeric: true });
});

// メタ文書が指す世代だけを読み込み、保存途中のチャンクや旧世代の混在を防ぐ。
// generationId がない従来データは chunk_0, chunk_1... をそのまま読み込める。
export const selectSalesChunkDocuments = (documents = [], meta = null) => {
  const normalizedDocuments = Array.isArray(documents) ? documents : [];
  const activeGenerationId = String(meta?.generationId || '');
  if (activeGenerationId) {
    return sortSalesChunkDocuments(normalizedDocuments.filter((snapshotDoc) => (
      String(readSalesChunkDocumentData(snapshotDoc).generationId || '') === activeGenerationId
    )));
  }

  const legacyDocuments = normalizedDocuments.filter((snapshotDoc) => {
    const data = readSalesChunkDocumentData(snapshotDoc);
    return !data.generationId && /^chunk_\d+$/.test(String(snapshotDoc?.id || ''));
  });
  if (legacyDocuments.length > 0) return sortSalesChunkDocuments(legacyDocuments);

  // 世代データはメタ文書による確定前には採用しない。最新らしいIDの推測は保存途中の世代を拾う恐れがある。
  return [];
};

export const isSalesChunkSelectionComplete = (selectedDocuments = [], meta = null) => {
  const expectedChunkCount = Number(meta?.chunkCount);
  if (!meta?.generationId || !Number.isInteger(expectedChunkCount) || expectedChunkCount < 0) return true;
  return selectedDocuments.length === expectedChunkCount;
};
