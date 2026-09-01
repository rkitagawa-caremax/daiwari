import { parseCSVLine } from '../lib/csv.js';
import {
  canPlacePanelAt,
  findFirstPlaceableIndex,
  getSpansFromSizeTypeRobust
} from './panelLayout.js';

export const DEFAULT_PDF_GRID_BOUNDS = Object.freeze({
  left: 10.4,
  top: 6,
  right: 11.8,
  bottom: 8.7
});

export const PDF_CROP_SIZE_OPTIONS = Object.freeze([
  '1/16（1コマ）',
  '1/8 縦（2コマ）',
  '1/8 横（2コマ）',
  '3/16 縦（3コマ）',
  '3/16 横（3コマ）',
  '1/4（4コマ）',
  '1/4 縦（4コマ）',
  '1/4 横（4コマ）',
  '6/16 縦（6コマ）',
  '6/16 横（6コマ）',
  '1/2 縦（8コマ）',
  '1/2 横（8コマ）',
  '12/16 縦（12コマ）',
  '12/16 横（12コマ）',
  '1P（16コマ）'
]);

const normalizeHeader = (value = '') => String(value)
  .normalize('NFKC')
  .trim()
  .toLowerCase()
  .replace(/[\s_-]+/g, '');

const HEADER_ALIASES = Object.freeze({
  pageNumber: ['ページ数', 'ページ', 'page', 'pageno', '頁'],
  order: ['追番', '順番', 'order'],
  frameNumber: ['コマ番号', '枠番号', 'frameno'],
  code: ['介援隊コード', '介援隊cd', '商品コード', 'code'],
  sizeType: ['コマ数', 'コマサイズ', 'サイズ', 'sizetype'],
  text: ['テキスト情報', 'テキスト', 'text'],
  coordinate: ['座標', 'coordinate', 'position'],
  xPos: ['xpos', 'x座標'],
  yPos: ['ypos', 'y座標']
});

const findHeaderIndex = (headers, aliases, fallback) => {
  const normalizedAliases = new Set(aliases.map(normalizeHeader));
  const index = headers.findIndex((header) => normalizedAliases.has(normalizeHeader(header)));
  return index >= 0 ? index : fallback;
};

const splitCsvRecords = (text = '') => {
  const records = [];
  let current = '';
  let inQuotes = false;

  for (let index = 0; index < text.length; index++) {
    const character = text[index];
    if (character === '"') {
      current += character;
      if (inQuotes && text[index + 1] === '"') {
        current += text[index + 1];
        index++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if ((character === '\n' || character === '\r') && !inQuotes) {
      if (character === '\r' && text[index + 1] === '\n') index++;
      if (current.trim()) records.push(current);
      current = '';
      continue;
    }
    current += character;
  }

  if (current.trim()) records.push(current);
  return records;
};

const clampGridPosition = (value) => {
  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) && parsed >= 1 && parsed <= 4 ? parsed : null;
};

const parseCoordinate = (value = '') => {
  const match = String(value).normalize('NFKC').match(/X\s*([1-4])\s*Y\s*([1-4])/i);
  return match ? { xPos: Number(match[1]), yPos: Number(match[2]) } : { xPos: null, yPos: null };
};

export const normalizePdfCropCode = (value = '') => {
  const normalized = String(value).normalize('NFKC').trim().toUpperCase();
  const matches = normalized.match(/[A-Z]{1,2}\d{3,5}/g);
  if (matches?.length) return matches[matches.length - 1];
  return '';
};

export const inferCatalogStartPage = (filename = '') => {
  const normalized = String(filename).normalize('NFKC');
  const match = normalized.match(/(?:^|[^A-Z0-9])P\s*0*(\d{1,4})(?:[^0-9]|$)/i);
  return match ? Math.max(1, Number.parseInt(match[1], 10)) : 1;
};

const isSupportedSizeType = (value = '') => {
  const normalized = String(value).normalize('NFKC').toLowerCase().replace(/\s+/g, '');
  return /(1\/16|1\/8|3\/16|1\/4|6\/16|1\/2|12\/16|1p|(?:1|2|3|4|6|8|12|16)コマ)/.test(normalized);
};

const occupy = (occupied, startIndex, rowSpan, colSpan) => {
  const startRow = Math.floor(startIndex / 4);
  const startCol = startIndex % 4;
  for (let row = 0; row < rowSpan; row++) {
    for (let col = 0; col < colSpan; col++) {
      occupied.add((startRow + row) * 4 + startCol + col);
    }
  }
};

const resolveMissingPositions = (rows, issues) => {
  const groups = new Map();
  rows.forEach((row) => {
    if (!groups.has(row.pageNumber)) groups.set(row.pageNumber, []);
    groups.get(row.pageNumber).push(row);
  });

  groups.forEach((pageRows, pageNumber) => {
    const occupied = new Set();
    const sortedRows = [...pageRows].sort((left, right) => (
      (left.frameNumber || left.order || left.csvRow) - (right.frameNumber || right.order || right.csvRow)
    ));

    sortedRows.forEach((row) => {
      let startIndex = null;
      if (row.xPos && row.yPos) {
        startIndex = (row.yPos - 1) * 4 + row.xPos - 1;
        if (!canPlacePanelAt(startIndex, row.rowSpan, row.colSpan, occupied)) {
          row.layoutStatus = 'conflict';
          issues.push({
            type: 'layout-conflict',
            csvRow: row.csvRow,
            message: `CSV ${row.csvRow}行目（${row.code}）の座標またはサイズが他のコマと重なっています。`
          });
          return;
        }
        row.positionSource = 'csv';
      } else {
        const orderHint = row.order > 0 && row.order <= 16 ? row.order - 1 : 0;
        startIndex = findFirstPlaceableIndex(row.rowSpan, row.colSpan, occupied, orderHint);
        if (startIndex < 0) {
          row.layoutStatus = 'unresolved';
          issues.push({
            type: 'layout-unresolved',
            csvRow: row.csvRow,
            message: `CSV ${row.csvRow}行目（${row.code}）の配置位置を決定できません。`
          });
          return;
        }
        row.xPos = (startIndex % 4) + 1;
        row.yPos = Math.floor(startIndex / 4) + 1;
        row.positionSource = 'estimated';
      }
      occupy(occupied, startIndex, row.rowSpan, row.colSpan);
      row.layoutStatus = 'ready';
    });

    if (pageRows.length === 0) {
      issues.push({ type: 'empty-page', pageNumber, message: `Page ${pageNumber} に切り抜き対象がありません。` });
    }
  });
};

export const parsePdfCropCsv = (content, { parseLine = parseCSVLine } = {}) => {
  const records = splitCsvRecords(String(content || '').replace(/^\uFEFF/, ''));
  if (records.length === 0) return { rows: [], issues: [{ type: 'empty-csv', message: 'CSVにデータがありません。' }] };

  const headers = parseLine(records[0]);
  const indexes = {
    pageNumber: findHeaderIndex(headers, HEADER_ALIASES.pageNumber, 1),
    order: findHeaderIndex(headers, HEADER_ALIASES.order, 2),
    frameNumber: findHeaderIndex(headers, HEADER_ALIASES.frameNumber, 3),
    code: findHeaderIndex(headers, HEADER_ALIASES.code, 4),
    sizeType: findHeaderIndex(headers, HEADER_ALIASES.sizeType, 5),
    text: findHeaderIndex(headers, HEADER_ALIASES.text, 7),
    coordinate: findHeaderIndex(headers, HEADER_ALIASES.coordinate, 8),
    xPos: findHeaderIndex(headers, HEADER_ALIASES.xPos, 10),
    yPos: findHeaderIndex(headers, HEADER_ALIASES.yPos, 11)
  };

  const rows = [];
  const issues = [];

  records.slice(1).forEach((record, recordIndex) => {
    const values = parseLine(record);
    const csvRow = recordIndex + 2;
    const rawCode = values[indexes.code] || '';
    const rawText = values[indexes.text] || '';
    if (!rawCode || rawCode === 'ダミーコマ' || rawText.trim()) return;

    const code = normalizePdfCropCode(rawCode);
    if (!code) {
      issues.push({ type: 'invalid-code', csvRow, message: `CSV ${csvRow}行目の介援隊コードを認識できません。` });
      return;
    }

    const pageNumber = Number.parseInt(values[indexes.pageNumber], 10);
    if (!Number.isInteger(pageNumber) || pageNumber < 1) {
      issues.push({ type: 'invalid-page', csvRow, message: `CSV ${csvRow}行目（${code}）のページ番号が不正です。` });
      return;
    }

    const sizeType = (values[indexes.sizeType] || '').trim();
    if (!isSupportedSizeType(sizeType)) {
      issues.push({ type: 'invalid-size', csvRow, message: `CSV ${csvRow}行目（${code}）のコマサイズを認識できません。` });
      return;
    }
    const { r: rowSpan, c: colSpan } = getSpansFromSizeTypeRobust(sizeType);
    const coordinate = parseCoordinate(values[indexes.coordinate]);
    const xPos = clampGridPosition(values[indexes.xPos]) || coordinate.xPos;
    const yPos = clampGridPosition(values[indexes.yPos]) || coordinate.yPos;
    const order = Number.parseInt(values[indexes.order], 10) || 0;
    const frameNumber = Number.parseInt(values[indexes.frameNumber], 10) || 0;

    rows.push({
      id: `${pageNumber}-${code}-${csvRow}`,
      csvRow,
      pageNumber,
      order,
      frameNumber,
      rawCode,
      code,
      filename: `${code}.jpg`,
      sizeType,
      rowSpan,
      colSpan,
      xPos,
      yPos,
      positionSource: xPos && yPos ? 'csv' : 'estimated',
      layoutStatus: 'pending'
    });
  });

  resolveMissingPositions(rows, issues);
  return { rows, issues };
};

export const updatePdfCropRowSize = (row, sizeType) => {
  const { r: rowSpan, c: colSpan } = getSpansFromSizeTypeRobust(sizeType);
  return { ...row, sizeType, rowSpan, colSpan };
};

export const getPdfCropRect = (row, bounds = DEFAULT_PDF_GRID_BOUNDS) => {
  const left = Number(bounds.left) / 100;
  const top = Number(bounds.top) / 100;
  const usableWidth = 1 - left - Number(bounds.right) / 100;
  const usableHeight = 1 - top - Number(bounds.bottom) / 100;
  const cellWidth = usableWidth / 4;
  const cellHeight = usableHeight / 4;
  return {
    x: left + (row.xPos - 1) * cellWidth,
    y: top + (row.yPos - 1) * cellHeight,
    width: row.colSpan * cellWidth,
    height: row.rowSpan * cellHeight
  };
};

export const isPdfCropRowInsideGrid = (row) => (
  row.layoutStatus !== 'conflict'
  && Number.isInteger(row.xPos)
  && Number.isInteger(row.yPos)
  && row.xPos >= 1
  && row.yPos >= 1
  && row.xPos + row.colSpan - 1 <= 4
  && row.yPos + row.rowSpan - 1 <= 4
);

export const pdfTextItemsContainCode = (textItems, row, bounds = DEFAULT_PDF_GRID_BOUNDS) => {
  if (!row.code || !Array.isArray(textItems)) return false;
  const rect = getPdfCropRect(row, bounds);
  return textItems.some((item) => (
    item.x >= rect.x
    && item.x <= rect.x + rect.width
    && item.y >= rect.y
    && item.y <= rect.y + rect.height
    && normalizePdfCropCode(item.text) === row.code
  ));
};
