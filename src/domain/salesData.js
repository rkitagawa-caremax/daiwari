import { parseCSVLine } from '../lib/csv.js';
import { normalizeCode } from './productCodes.js';

export const SALES_DATA_HEADER_ROW_COUNT = 2;
export const SALES_DATA_CHUNK_SIZE = 1000;
export const SALES_DATA_CHUNK_MAX_BYTES = 640 * 1024;
export const SALES_DATA_WRITE_BATCH_SIZE = 5;

const normalizeSalesHeader = (value) => String(value ?? '')
  .normalize('NFKC')
  .replace(/[\s\u3000]/g, '');

const parseSalesCount = (value) => {
  const normalized = String(value ?? '')
    .normalize('NFKC')
    .replace(/[,，\s]/g, '');
  if (!normalized) return 0;
  const negativeMatch = normalized.match(/^\(([-+]?\d+(?:\.\d+)?)\)$/);
  const parsed = Number(negativeMatch ? `-${negativeMatch[1]}` : normalized);
  return Number.isFinite(parsed) ? parsed : (parseInt(normalized, 10) || 0);
};

export const resolveSalesMonthColumns = (headerRows = []) => {
  const rows = (Array.isArray(headerRows) ? headerRows : []).map((row) => (
    Array.isArray(row) ? row : []
  ));
  const columnCount = Math.max(0, ...rows.map((row) => row.length));
  const columns = [];
  let carriedYear = null;
  let previousMonth = null;

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    // 現行帳票の0〜3列は商品情報。帳票タイトル内の年月を売上列と誤認しない。
    if (columnIndex <= 3) continue;
    const headerParts = rows
      .map((row) => normalizeSalesHeader(row[columnIndex]))
      .filter(Boolean);
    const header = headerParts.join(' ');
    const explicitYear = header.match(/(?:19|20)\d{2}/)?.[0];
    if (explicitYear) carriedYear = Number(explicitYear);
    if (/合計|累計|平均|前年差|前年比/.test(header)) continue;

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

// 売上CSVのUIや保存先に依存しない解析処理。
export const parseSalesCsvContent = (
  csvText,
  { headerRowCount = SALES_DATA_HEADER_ROW_COUNT } = {}
) => {
  const normalizedText = String(csvText ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rows = normalizedText.split('\n');
  const parsedHeaderRows = rows.slice(0, headerRowCount).map((row) => parseCSVLine(row));
  const monthColumns = resolveSalesMonthColumns(parsedHeaderRows);
  const salesData = {};

  rows.slice(headerRowCount).forEach((row) => {
    if (!row.trim()) return;
    const columns = parseCSVLine(row);
    if (columns.length <= 17) return;

    const rawCode = columns[3];
    if (!rawCode) return;

    const code = normalizeCode(rawCode);
    const count = parseSalesCount(columns[17]);
    const monthlySales = monthColumns.map((monthColumn) => parseSalesCount(columns[monthColumn.columnIndex]));
    if (!salesData[code]) salesData[code] = [];
    const item = {
      name: columns[1] || '',
      spec: columns[2] || '',
      count
    };
    if (monthlySales.length > 0) {
      item.monthlySales = monthlySales;
      // 月ラベルはコードごとに1度だけ保存し、Firestoreの通信量を抑える。
      if (salesData[code].length === 0) item.monthlyLabels = monthColumns.map((monthColumn) => monthColumn.label);
    }
    salesData[code].push(item);
  });

  return salesData;
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
