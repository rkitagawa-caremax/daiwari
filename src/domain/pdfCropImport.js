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
  pageNumber: ['ページ数', 'ページ番号', 'ページ', 'page', 'pageno', '頁'],
  order: ['追番', '順番', 'order'],
  frameNumber: ['コマ番号', '枠番号', 'frameno'],
  code: ['介援隊コード', '介援隊cd', '商品コード', 'code'],
  sizeType: ['コマ数', 'コマサイズ', 'サイズ', 'sizetype'],
  kind: ['コマ種別', '種別', 'kind'],
  text: ['テキスト情報', 'テキスト', 'text'],
  catalogName: ['掲載名', '商品名', 'catalogname', 'productname'],
  coordinate: ['座標', 'coordinate', 'position'],
  xPos: ['xpos', 'x座標'],
  yPos: ['ypos', 'y座標']
});

// 見出しが読めない CSV 用の列位置 (旧フォーマット: ジャンル,ページ数,追番,コマ番号,介援隊コード,コマ数,,テキスト情報,座標,コマID,X_POS,Y_POS)
const HEADER_FALLBACK_INDEXES = Object.freeze({
  pageNumber: 1,
  order: 2,
  frameNumber: 3,
  code: 4,
  sizeType: 5,
  text: 7,
  coordinate: 8,
  xPos: 10,
  yPos: 11
});

const REQUIRED_HEADER_KEYS = Object.freeze(['pageNumber', 'code', 'sizeType']);

const findHeaderIndex = (headers, aliases) => {
  const normalizedAliases = new Set(aliases.map(normalizeHeader));
  return headers.findIndex((header) => normalizedAliases.has(normalizeHeader(header)));
};

// 見出し行から列位置を決める。
// 主要な列 (ページ / 介援隊コード / コマサイズ) が見出しで揃っていれば、見つからない列は「無し」(-1) にする。
// 列の並びが違う CSV でも、余った列を別の意味に取り違えないようにするため。
// 見出しが揃っていない CSV だけ、旧フォーマットの列位置で読む。
export const resolvePdfCropCsvColumns = (headers = []) => {
  const matched = Object.fromEntries(
    Object.entries(HEADER_ALIASES).map(([key, aliases]) => [key, findHeaderIndex(headers, aliases)])
  );
  if (REQUIRED_HEADER_KEYS.every((key) => matched[key] >= 0)) return matched;
  return Object.fromEntries(Object.entries(matched).map(([key, index]) => [
    key,
    index >= 0 ? index : (HEADER_FALLBACK_INDEXES[key] ?? -1)
  ]));
};

const valueAt = (values, index) => (index >= 0 ? String(values[index] ?? '') : '');

// コマ種別が「タイトル」「見出し」などの行は切り抜き対象ではない
const NON_PRODUCT_KIND = /^(タイトル|見出し|ダミー)/;

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
    // 座標が無い行を追番順に置けるのは「そのページが全く座標を持たず、4x4 に収まる」ときだけ。
    // 一部の行だけ座標が空欄なら CSV 上の意図的な空欄 (グリッド外のコマ) なので推測しない。
    // 部品表のようにコマ数がページを超えるページも同様。
    const canEstimatePositions = sortedRows.every((row) => !(row.xPos && row.yPos))
      && sortedRows.reduce((cells, row) => cells + row.rowSpan * row.colSpan, 0) <= 16;

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
      } else if (!canEstimatePositions) {
        row.layoutStatus = 'unresolved';
        issues.push({
          type: 'layout-unresolved',
          csvRow: row.csvRow,
          message: `CSV ${row.csvRow}行目（${row.code}）は座標が空欄のため配置できません。`
        });
        return;
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
  const indexes = resolvePdfCropCsvColumns(headers);

  const rows = [];
  const issues = [];

  records.slice(1).forEach((record, recordIndex) => {
    const values = parseLine(record);
    const csvRow = recordIndex + 2;
    const rawCode = valueAt(values, indexes.code);
    const rawText = valueAt(values, indexes.text);
    const kind = valueAt(values, indexes.kind).trim();
    if (!rawCode || rawCode === 'ダミーコマ' || rawText.trim() || NON_PRODUCT_KIND.test(kind)) return;

    const code = normalizePdfCropCode(rawCode);
    if (!code) {
      issues.push({ type: 'invalid-code', csvRow, message: `CSV ${csvRow}行目の介援隊コードを認識できません。` });
      return;
    }

    const pageNumber = Number.parseInt(valueAt(values, indexes.pageNumber), 10);
    if (!Number.isInteger(pageNumber) || pageNumber < 1) {
      issues.push({ type: 'invalid-page', csvRow, message: `CSV ${csvRow}行目（${code}）のページ番号が不正です。` });
      return;
    }

    const sizeType = valueAt(values, indexes.sizeType).trim();
    if (!isSupportedSizeType(sizeType)) {
      issues.push({ type: 'invalid-size', csvRow, message: `CSV ${csvRow}行目（${code}）のコマサイズを認識できません。` });
      return;
    }
    const { r: rowSpan, c: colSpan } = getSpansFromSizeTypeRobust(sizeType);
    const coordinate = parseCoordinate(valueAt(values, indexes.coordinate));
    const xPos = clampGridPosition(valueAt(values, indexes.xPos)) || coordinate.xPos;
    const yPos = clampGridPosition(valueAt(values, indexes.yPos)) || coordinate.yPos;
    const order = Number.parseInt(valueAt(values, indexes.order), 10) || 0;
    const frameNumber = Number.parseInt(valueAt(values, indexes.frameNumber), 10) || 0;
    const catalogName = valueAt(values, indexes.catalogName).normalize('NFKC').trim();

    rows.push({
      id: `${pageNumber}-${code}-${csvRow}`,
      csvRow,
      pageNumber,
      order,
      frameNumber,
      rawCode,
      code,
      filename: `${code}.jpg`,
      ...(catalogName ? { catalogName } : {}),
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

// 余白設定 (%) から 4x4 グリッドの原点とセルサイズ (正規化座標) を求める
export const getPdfGridGeometry = (bounds = DEFAULT_PDF_GRID_BOUNDS) => {
  const left = Number(bounds.left) / 100;
  const top = Number(bounds.top) / 100;
  const usableWidth = 1 - left - Number(bounds.right) / 100;
  const usableHeight = 1 - top - Number(bounds.bottom) / 100;
  return { left, top, cellWidth: usableWidth / 4, cellHeight: usableHeight / 4 };
};

export const getPdfCropRectFromGrid = (row, grid) => ({
  x: grid.left + (row.xPos - 1) * grid.cellWidth,
  y: grid.top + (row.yPos - 1) * grid.cellHeight,
  width: row.colSpan * grid.cellWidth,
  height: row.rowSpan * grid.cellHeight
});

export const getPdfCropRect = (row, bounds = DEFAULT_PDF_GRID_BOUNDS) => (
  getPdfCropRectFromGrid(row, getPdfGridGeometry(bounds))
);

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
export const PDF_CROP_TEXT_VERSION = 2;

const compactPdfText = (values) => values
  .map((value) => String(value || '').normalize('NFKC').trim())
  .filter(Boolean)
  .join(' ')
  .replace(/\s+/g, ' ')
  .replace(/([\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}])\s+(?=[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}])/gu, '$1')
  .trim();

const getPdfTextItemCenter = (item) => ({
  x: Number(item?.x || 0) + Math.max(0, Number(item?.width || 0)) / 2,
  y: Number(item?.y || 0) - Math.max(0, Number(item?.height || 0)) * 0.34
});

const isTextItemCenteredInRect = (item, rect) => {
  const center = getPdfTextItemCenter(item);
  return center.x >= rect.x
    && center.x <= rect.x + rect.width
    && center.y >= rect.y
    && center.y <= rect.y + rect.height;
};

const sortPdfTextItemsForReading = (items = []) => {
  const positioned = items.map((item) => ({
    item,
    center: getPdfTextItemCenter(item),
    height: Math.max(0, Number(item?.height || 0))
  })).sort((left, right) => left.center.y - right.center.y || left.center.x - right.center.x);
  const lines = [];

  positioned.forEach((entry) => {
    const currentLine = lines[lines.length - 1];
    const tolerance = Math.max(0.003, Math.max(entry.height, currentLine?.maxHeight || 0) * 0.65);
    if (!currentLine || Math.abs(entry.center.y - currentLine.y) > tolerance) {
      lines.push({ y: entry.center.y, maxHeight: entry.height, entries: [entry] });
      return;
    }
    currentLine.entries.push(entry);
    currentLine.y = currentLine.entries.reduce((sum, value) => sum + value.center.y, 0) / currentLine.entries.length;
    currentLine.maxHeight = Math.max(currentLine.maxHeight, entry.height);
  });

  return lines.flatMap((line) => line.entries
    .sort((left, right) => left.center.x - right.center.x)
    .map((entry) => entry.item.text));
};

const intersectPdfRects = (left, right) => {
  if (!left) return right || null;
  if (!right) return left;
  const x = Math.max(left.x, right.x);
  const y = Math.max(left.y, right.y);
  const rightEdge = Math.min(left.x + left.width, right.x + right.width);
  const bottomEdge = Math.min(left.y + left.height, right.y + right.height);
  if (rightEdge <= x || bottomEdge <= y) return null;
  return { x, y, width: rightEdge - x, height: bottomEdge - y };
};

// 自動枠が隣接コマ側へ広がっても、文字はCSVグリッドと文字アンカーの内側だけから取得する。
// 手動枠はユーザーが確定した範囲をそのまま尊重する。
export const resolvePdfTextExtractionRect = ({ cropRect, gridRect, textRect, isManual = false } = {}) => {
  if (isManual) return cropRect || textRect || gridRect || null;
  let resolved = intersectPdfRects(cropRect, gridRect) || gridRect || cropRect || null;
  if (textRect) resolved = intersectPdfRects(resolved, textRect) || resolved;
  return resolved;
};

// 指定矩形 (正規化座標) 内の文字を保存用に整形して返す
export const extractPdfTextInRect = (textItems, rect, maxLength = MAX_PDF_CROP_TEXT_LENGTH) => {
  if (!Array.isArray(textItems) || !rect) return { text: '', truncated: false };
  const compacted = compactPdfText(sortPdfTextItemsForReading(textItems
    .filter((item) => isTextItemCenteredInRect(item, rect))));
  const safeMaxLength = Math.max(0, Number.parseInt(maxLength, 10) || 0);
  if (!safeMaxLength || compacted.length <= safeMaxLength) {
    return { text: compacted, truncated: false };
  }
  return { text: compacted.slice(0, safeMaxLength), truncated: true };
};

const findPrimaryPrice = (candidates, referenceIndex) => {
  if (candidates.length === 0) return null;
  return [...candidates].sort((left, right) => {
    if (referenceIndex < 0) return left.index - right.index;
    const leftDistance = Math.abs(left.index - referenceIndex) + (left.index < referenceIndex ? 10000 : 0);
    const rightDistance = Math.abs(right.index - referenceIndex) + (right.index < referenceIndex ? 10000 : 0);
    return leftDistance - rightDistance;
  })[0];
};

// 金額を単一値に決め打ちせず候補も残す。将来の検索・突合では primary と候補の両方を利用できる。
export const extractPdfPriceFields = (sourceText = '', code = '') => {
  const normalizedText = String(sourceText || '').normalize('NFKC');
  const normalizedCode = normalizePdfCropCode(code);
  const candidates = [];
  const seen = new Set();
  const pricePattern = /[¥￥]\s*([0-9][0-9,]*)/g;
  let match;
  while ((match = pricePattern.exec(normalizedText)) !== null) {
    const amount = Number.parseInt(match[1].replace(/,/g, ''), 10);
    if (!Number.isFinite(amount)) continue;
    const prefix = normalizedText.slice(Math.max(0, match.index - 12), match.index);
    const taxType = /税\s*抜/.test(prefix) ? 'excluding' : 'including';
    const key = `${taxType}:${amount}`;
    if (seen.has(key)) continue;
    seen.add(key);
    candidates.push({ amount, taxType, text: match[0].replace(/\s+/g, ''), index: match.index });
  }

  const codeIndex = normalizedCode ? normalizedText.toUpperCase().indexOf(normalizedCode) : -1;
  const including = findPrimaryPrice(candidates.filter((candidate) => candidate.taxType === 'including'), codeIndex);
  const excludingReference = including?.index ?? codeIndex;
  const excluding = findPrimaryPrice(candidates.filter((candidate) => candidate.taxType === 'excluding'), excludingReference);
  const compactCandidates = candidates.slice(0, 20).map((candidate) => ({
    amount: candidate.amount,
    taxType: candidate.taxType,
    text: candidate.text
  }));

  return {
    ...(including ? { priceIncludingTax: including.amount } : {}),
    ...(excluding ? { priceExcludingTax: excluding.amount } : {}),
    ...(compactCandidates.length > 0 ? { priceCandidates: compactCandidates } : {}),
    priceExtractionConfidence: including && codeIndex >= 0 && including.index >= codeIndex && including.index - codeIndex <= 200
      ? 'high'
      : (including ? 'medium' : 'none')
  };
};

const uniqueStrings = (values, limit) => [...new Set(values.filter(Boolean))].slice(0, limit);

const extractHandlingMarkers = (text) => uniqueStrings(
  [...text.matchAll(/[（(]\s*([A-Z]{1,3})\s*[）)]/gi)].map((match) => `(${match[1].toUpperCase()})`),
  20
);

const extractSpecificationDetails = (text) => {
  const specifications = uniqueStrings(text
    .split('●')
    .slice(1)
    .map((value) => value
      .split(/(?=在庫商品|直送商品|メーカー直送|受注生産|返品不可|株式会社|有限会社|\(株\))/)[0]
      .trim())
    .filter(Boolean)
    .map((value) => `●${value.slice(0, 300)}`), 30);
  return {
    specifications,
    compositionDetails: specifications.filter((value) => /成分|原材料|栄養/.test(value)).slice(0, 12),
    materialDetails: specifications.filter((value) => /材質|素材/.test(value)).slice(0, 12)
  };
};

const extractCatchCopyCandidates = (text, productName, code) => {
  const normalizedProductName = String(productName || '').normalize('NFKC').trim();
  const normalizedCode = normalizePdfCropCode(code);
  const chunks = [];
  let current = '';
  for (const character of text) {
    current += character;
    if (/[。！？!?♪]/.test(character) || current.length >= 180) {
      chunks.push(current);
      current = '';
    }
  }
  if (current) chunks.push(current);

  const proofBoilerplate = /校正|変更あり|変更なし|ご返答|メール到着|カタログ紙面|商品情報|メーカーご担当|営業企画部|アップをお願い|JANコード|CMCystem/i;
  return uniqueStrings(chunks.map((value) => {
    let candidate = value.replace(/\s+/g, ' ').trim();
    if (normalizedProductName) candidate = candidate.replace(normalizedProductName, '').trim();
    if (normalizedCode && candidate.toUpperCase().includes(normalizedCode)) {
      candidate = candidate.slice(candidate.toUpperCase().lastIndexOf(normalizedCode) + normalizedCode.length).trim();
    }
    if (/[¥￥]/.test(candidate)) {
      candidate = candidate.replace(/^.*[¥￥]\s*[0-9][0-9,]*\s*[）)]?\s*/, '').trim();
    }
    candidate = candidate.replace(/^\d{1,3}\s+/, '').trim();
    return candidate;
  }).filter((candidate) => (
    candidate.length >= 8
    && candidate.length <= 180
    && /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u.test(candidate)
    && !proofBoilerplate.test(candidate)
    && !candidate.includes('●')
    && !/[¥￥]/.test(candidate)
    && !(normalizedCode && candidate.toUpperCase().includes(normalizedCode))
    && !/税抜|生産国|材質|成分|原材料/.test(candidate)
  )), 8);
};

const extractItemNumberCandidates = (text, code) => {
  const normalizedCode = normalizePdfCropCode(code);
  const codeIndex = normalizedCode ? text.toUpperCase().indexOf(normalizedCode) : -1;
  const matches = [];
  const addMatches = (pattern, confidence) => {
    for (const match of text.matchAll(pattern)) {
      const value = String(match[0] || '').toUpperCase();
      if (!value || value === normalizedCode || value.includes(`-${normalizedCode}`)) continue;
      if (/^\d{1,2}-\d{1,2}$/.test(value)) continue;
      matches.push({ value, index: match.index ?? 0, confidence });
    }
  };

  addMatches(/(?<![A-Z0-9])[A-Z0-9]{2,10}-[A-Z0-9-]{2,14}(?![A-Z0-9])/gi, 'high');
  addMatches(/(?<![A-Z0-9])(?=[A-Z0-9]{5,18}(?![A-Z0-9]))(?=[A-Z0-9]*[A-Z])(?=[A-Z0-9]*\d)[A-Z0-9]{5,18}(?![A-Z0-9])/gi, 'medium');

  if (codeIndex >= 0) {
    const nearbyStart = Math.max(0, codeIndex - 60);
    const nearby = text.slice(nearbyStart, Math.min(text.length, codeIndex + normalizedCode.length + 100));
    for (const match of nearby.matchAll(/(?<![A-Z0-9,])\d{4,8}(?![A-Z0-9,])/gi)) {
      const index = nearbyStart + (match.index ?? 0);
      const suffix = text.slice(index + match[0].length, index + match[0].length + 5);
      const prefix = text.slice(Math.max(0, index - 2), index);
      if (/^(?:g|kg|mg|mL|L|cm|mm|枚|本|個|食|円|kcal)/i.test(suffix.trim())) continue;
      if (/[¥￥]/.test(prefix)) continue;
      matches.push({ value: match[0], index, confidence: 'medium' });
    }
  }

  const candidates = [];
  const seen = new Set();
  matches.sort((left, right) => {
    if (codeIndex < 0) return left.index - right.index;
    return Math.abs(left.index - codeIndex) - Math.abs(right.index - codeIndex);
  }).forEach((match) => {
    if (seen.has(match.value)) return;
    seen.add(match.value);
    candidates.push(match);
  });
  const limited = candidates.slice(0, 20);
  return {
    ...(limited[0] ? {
      itemNumber: limited[0].value,
      itemNumberExtractionConfidence: limited[0].confidence
    } : {}),
    itemNumberCandidates: limited.map((candidate) => candidate.value)
  };
};

export const extractPdfCatalogDetails = (sourceText = '', { code = '', productName = '' } = {}) => {
  const text = String(sourceText || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const handlingMarkers = extractHandlingMarkers(text);
  const availabilityLabels = uniqueStrings([
    ...(/在庫商品/.test(text) ? ['在庫商品'] : []),
    ...(/(?:メーカー直送|直送商品|直送品|直送)/.test(text) ? ['直送'] : []),
    ...(/受注生産/.test(text) ? ['受注生産'] : []),
    ...(/返品不可/.test(text) ? ['返品不可'] : []),
    ...(/ケース販売/.test(text) ? ['ケース販売'] : [])
  ], 10);
  const specificationData = extractSpecificationDetails(text);
  const catchCopyCandidates = extractCatchCopyCandidates(text, productName, code);
  const itemNumberData = extractItemNumberCandidates(text, code);
  const hasStock = availabilityLabels.includes('在庫商品');
  const hasDirect = availabilityLabels.includes('直送');

  return {
    version: 1,
    ...itemNumberData,
    ...(catchCopyCandidates[0] ? {
      catchCopy: catchCopyCandidates[0],
      catchCopyExtractionConfidence: 'medium'
    } : {}),
    catchCopyCandidates,
    availability: hasStock && hasDirect ? 'mixed' : (hasStock ? 'stock' : (hasDirect ? 'direct' : 'unknown')),
    availabilityLabels,
    handlingMarkers,
    hasDemoMarker: handlingMarkers.includes('(D)') || /デモ機/.test(text),
    ...specificationData
  };
};

export const extractPdfCatalogTextData = ({ textItems = [], rect, code = '', catalogName = '' } = {}) => {
  const sourceText = extractPdfTextInRect(textItems, rect);
  const productName = String(catalogName || '').normalize('NFKC').trim();
  return {
    sourceText: sourceText.text,
    sourceTextTruncated: sourceText.truncated,
    sourceTextVersion: PDF_CROP_TEXT_VERSION,
    catalogCode: normalizePdfCropCode(code),
    productName,
    productNameSource: productName ? 'csv' : 'unavailable',
    ...extractPdfPriceFields(sourceText.text, code),
    catalogTextData: extractPdfCatalogDetails(sourceText.text, { code, productName })
  };
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
