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

export const MAX_PDF_CROP_BATCH_PAGES = 10;

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

export const buildPdfCropBatchPages = (sources = []) => sources.flatMap((source, sourceIndex) => {
  const numPages = Math.max(1, Number.parseInt(source?.numPages, 10) || 1);
  const catalogStartPage = inferCatalogStartPage(source?.file?.name || source?.name || '');
  return Array.from({ length: numPages }, (_, pageIndex) => ({
    id: `${sourceIndex + 1}-${pageIndex + 1}`,
    sourceIndex,
    file: source?.file || source,
    filename: source?.file?.name || source?.name || '',
    pdfPageNumber: pageIndex + 1,
    catalogPage: catalogStartPage + pageIndex
  }));
});

// 一括処理のページ一覧に、ユーザーが手動で変更した対象ページ (catalogPage) を反映する。
export const applyPdfCropCatalogPageOverrides = (batchPages = [], overrides = {}) => batchPages.map((page) => {
  const override = Number.parseInt(overrides?.[page.id], 10);
  return Number.isInteger(override) && override >= 1 ? { ...page, catalogPage: override } : page;
});

// 同一ページ内で座標が重なるコマ (およびグリッド外のコマ) の id を返す。
export const findPdfCropGridConflicts = (rows = []) => {
  const conflicts = new Set();
  const occupied = new Map();
  rows.forEach((row) => {
    if (!isPdfCropRowInsideGrid({ ...row, layoutStatus: 'ready' })) {
      conflicts.add(row.id);
      return;
    }
    for (let y = row.yPos; y < row.yPos + row.rowSpan; y++) {
      for (let x = row.xPos; x < row.xPos + row.colSpan; x++) {
        const key = `${x}:${y}`;
        if (occupied.has(key)) {
          conflicts.add(row.id);
          conflicts.add(occupied.get(key));
        } else {
          occupied.set(key, row.id);
        }
      }
    }
  });
  return conflicts;
};

const sortPdfCropRows = (rows) => [...rows].sort((left, right) => (
  (left.frameNumber || left.order || left.csvRow) - (right.frameNumber || right.order || right.csvRow)
));

const hasPdfCropPosition = (row) => Number.isInteger(row?.xPos) && Number.isInteger(row?.yPos) && row.xPos >= 1 && row.yPos >= 1;

// 一括処理の各ページについて「CSV 上の対象コマ」「実際に保存するコマ」を決める。
// 目的は「切り抜いてコード名で保存する」ことなので、除外は最小限にする:
//   - コードが読めない / 座標が全く無い行 (切り抜き位置が決まらない)
//   - 同じバッチ内で既に登場したコード (同名ファイルを二重に作らない)
//   - skipExistingCodes 指定時のみ、ライブラリに同じコードがある行
// 座標の重なり (conflict) は警告として数えるが、切り抜き自体は行う。
// 同じ対象ページに複数の PDF ページが割り当たっている場合は isDuplicateCatalogPage を立てる。
export const buildPdfCropPagePlans = ({
  batchPages = [],
  rows = [],
  existingCodes = new Set(),
  skipExistingCodes = false
} = {}) => {
  const seenCodes = new Set();
  const catalogPageCounts = new Map();
  batchPages.forEach((page) => {
    catalogPageCounts.set(page.catalogPage, (catalogPageCounts.get(page.catalogPage) || 0) + 1);
  });

  return batchPages.map((page) => {
    const targetRows = sortPdfCropRows(rows.filter((row) => row.pageNumber === page.catalogPage));
    const conflictIds = findPdfCropGridConflicts(targetRows);
    let existingCount = 0;
    let unplaceableCount = 0;
    let duplicateCodeCount = 0;
    const importRows = targetRows.filter((row) => {
      const code = normalizePdfCropCode(row.code);
      if (!code || !hasPdfCropPosition(row)) {
        unplaceableCount++;
        return false;
      }
      const isExisting = existingCodes.has(code);
      if (isExisting) existingCount++;
      if (isExisting && skipExistingCodes) return false;
      if (seenCodes.has(code)) {
        duplicateCodeCount++;
        return false;
      }
      seenCodes.add(code);
      return true;
    });
    return {
      page,
      targetRows,
      conflictIds,
      importRows,
      existingCount,
      unplaceableCount,
      duplicateCodeCount,
      skippedCount: Math.max(0, targetRows.length - importRows.length),
      hasCsvRows: targetRows.length > 0,
      isDuplicateCatalogPage: (catalogPageCounts.get(page.catalogPage) || 0) > 1
    };
  });
};

export const summarizePdfCropPagePlans = (plans = []) => plans.reduce((summary, plan) => ({
  pageCount: summary.pageCount + 1,
  targetCount: summary.targetCount + plan.targetRows.length,
  importCount: summary.importCount + plan.importRows.length,
  skippedCount: summary.skippedCount + plan.skippedCount,
  existingCount: summary.existingCount + plan.existingCount,
  unplaceableCount: summary.unplaceableCount + plan.unplaceableCount,
  duplicateCodeCount: summary.duplicateCodeCount + plan.duplicateCodeCount,
  conflictCount: summary.conflictCount + plan.conflictIds.size,
  pagesWithoutCsv: summary.pagesWithoutCsv + (plan.hasCsvRows ? 0 : 1),
  duplicateCatalogPages: summary.duplicateCatalogPages + (plan.isDuplicateCatalogPage ? 1 : 0)
}), {
  pageCount: 0,
  targetCount: 0,
  importCount: 0,
  skippedCount: 0,
  existingCount: 0,
  unplaceableCount: 0,
  duplicateCodeCount: 0,
  conflictCount: 0,
  pagesWithoutCsv: 0,
  duplicateCatalogPages: 0
});

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

const isTextItemInRect = (item, rect) => (
  item.x >= rect.x
  && item.x <= rect.x + rect.width
  && item.y >= rect.y
  && item.y <= rect.y + rect.height
);

// 指定矩形 (正規化座標) 内に、そのコードの文字があるか
export const pdfTextItemsContainCodeInRect = (textItems, code, rect) => (
  !!code
  && !!rect
  && Array.isArray(textItems)
  && textItems.some((item) => isTextItemInRect(item, rect) && normalizePdfCropCode(item.text) === code)
);

export const pdfTextItemsContainCode = (textItems, row, bounds = DEFAULT_PDF_GRID_BOUNDS) => {
  if (!row.code || !Array.isArray(textItems)) return false;
  return pdfTextItemsContainCodeInRect(textItems, row.code, getPdfCropRect(row, bounds));
};

export const MAX_PDF_CROP_TEXT_LENGTH = 4000;

const compactPdfText = (values) => values
  .map((value) => String(value || '').normalize('NFKC').trim())
  .filter(Boolean)
  .join(' ')
  .replace(/\s+/g, ' ')
  .replace(/([\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}])\s+(?=[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}])/gu, '$1')
  .trim();

// 指定矩形 (正規化座標) 内の文字を保存用に整形して返す
export const extractPdfTextInRect = (textItems, rect, maxLength = MAX_PDF_CROP_TEXT_LENGTH) => {
  if (!Array.isArray(textItems) || !rect) return { text: '', truncated: false };
  const compacted = compactPdfText(textItems
    .filter((item) => isTextItemInRect(item, rect))
    .map((item) => item.text));
  const safeMaxLength = Math.max(0, Number.parseInt(maxLength, 10) || 0);
  if (!safeMaxLength || compacted.length <= safeMaxLength) {
    return { text: compacted, truncated: false };
  }
  return { text: compacted.slice(0, safeMaxLength), truncated: true };
};

export const extractPdfCropText = (
  textItems,
  row,
  bounds = DEFAULT_PDF_GRID_BOUNDS,
  maxLength = MAX_PDF_CROP_TEXT_LENGTH
) => {
  if (!Array.isArray(textItems) || !isPdfCropRowInsideGrid(row)) {
    return { text: '', truncated: false };
  }
  return extractPdfTextInRect(textItems, getPdfCropRect(row, bounds), maxLength);
};
