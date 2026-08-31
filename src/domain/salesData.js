import { parseCSVLine } from '../lib/csv.js';
import { normalizeCode } from './productCodes.js';

export const SALES_DATA_HEADER_ROW_COUNT = 2;
export const SALES_DATA_CHUNK_SIZE = 1000;

// 売上CSVのUIや保存先に依存しない解析処理。
export const parseSalesCsvContent = (
  csvText,
  { headerRowCount = SALES_DATA_HEADER_ROW_COUNT } = {}
) => {
  const normalizedText = String(csvText ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rows = normalizedText.split('\n');
  const salesData = {};

  rows.slice(headerRowCount).forEach((row) => {
    if (!row.trim()) return;
    const columns = parseCSVLine(row);
    if (columns.length <= 17) return;

    const rawCode = columns[3];
    if (!rawCode) return;

    const code = normalizeCode(rawCode);
    const count = parseInt(columns[17].replace(/,/g, '')) || 0;
    if (!salesData[code]) salesData[code] = [];
    salesData[code].push({
      name: columns[1] || '',
      spec: columns[2] || '',
      count
    });
  });

  return salesData;
};

export const splitSalesDataIntoChunks = (
  salesData,
  chunkSize = SALES_DATA_CHUNK_SIZE
) => {
  const entries = Object.entries(salesData || {});
  const chunks = [];
  for (let index = 0; index < entries.length; index += chunkSize) {
    chunks.push(Object.fromEntries(entries.slice(index, index + chunkSize)));
  }
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
