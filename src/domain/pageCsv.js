import { getPanelCsvCode } from './panels.js';
import { getSizeType } from './panelLayout.js';

// 台割ページ CSV (出力側) のカラム定義。
// 「追番」と結合済みの「座標」(X{n}Y{m}) は内部データとして扱い、CSVには出力しない。
// 配置は X_POS / Y_POS の数値（1始まり、4×4 グリッド内）として出力する。
export const PAGE_CSV_HEADERS = Object.freeze([
  'ジャンル', '介援隊コード', 'ページ数', 'X_POS', 'Y_POS', 'コマ番号', 'コマ数', 'コマID', 'テキスト情報'
]);

export const EXCLUDED_ITEMS_CSV_HEADERS = Object.freeze(['介援隊コード', '画像名', 'ラベル', '登録日時']);

const CSV_BOM = '﻿';

const escapeCsvText = (value) => {
  const text = value || '';
  return /[,"\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

const joinCsvLines = (headers, rows) => CSV_BOM + [headers.join(','), ...rows].join('\n');

// 1ページ分のコマを CSV 行 (文字列) の配列にする。
export const buildPageCsvRowsForSheet = (sheet, { pageNum, genreLabel }) => {
  const rows = [];
  let frameCounter = 0;

  (sheet?.panels || []).forEach((panel, panelIndex) => {
    if (panel.hidden) return;
    frameCounter++;
    const codeVal = getPanelCsvCode(panel);
    const sizeVal = panel.sizeType || getSizeType(panel.rowSpan || 1, panel.colSpan || 1);
    const textVal = escapeCsvText(panel.text || '');

    const gridRow = Math.floor(panelIndex / 4) + 1; // 1始まり
    const gridCol = (panelIndex % 4) + 1;           // 1始まり
    const panelIdVal = panel.panelId || '';

    rows.push([
      genreLabel,
      codeVal,
      pageNum,
      gridCol,
      gridRow,
      frameCounter,
      sizeVal,
      panelIdVal,
      textVal
    ].join(','));
  });

  return rows;
};

// 全ページを CSV 文字列 (BOM 付き) にする。
export const buildPageCsvContent = ({ sheets = [], genres = [] } = {}) => {
  const rows = [];
  sheets.forEach((sheet, sheetIndex) => {
    const genreLabel = genres.find((g) => g.id === sheet.genre)?.label || '未設定';
    rows.push(...buildPageCsvRowsForSheet(sheet, { pageNum: sheetIndex + 1, genreLabel }));
  });
  return joinCsvLines(PAGE_CSV_HEADERS, rows);
};

const formatExcludedItemDate = (createdAt, now) => {
  if (createdAt?.toDate) return createdAt.toDate().toLocaleString();
  if (createdAt?.seconds) return new Date(createdAt.seconds * 1000).toLocaleString();
  return now.toLocaleString();
};

// 除外リストを CSV 文字列 (BOM 付き) にする。
export const buildExcludedItemsCsvContent = (excludedItems = [], { now = new Date() } = {}) => {
  const rows = excludedItems.map((item) => [
    item.code || '',
    item.originalName || '',
    item.label || '',
    formatExcludedItemDate(item.createdAt, now)
  ].join(','));
  return joinCsvLines(EXCLUDED_ITEMS_CSV_HEADERS, rows);
};

export const buildDatedCsvFilename = (prefix, date = new Date()) => (
  `${prefix}_${date.toISOString().slice(0, 10)}.csv`
);
