import test from 'node:test';
import assert from 'node:assert/strict';

import {
  PAGE_CSV_HEADERS,
  buildDatedCsvFilename,
  buildExcludedItemsCsvContent,
  buildPageCsvContent,
  buildPageCsvRowsForSheet
} from '../src/domain/pageCsv.js';
import { buildDefaultPanels } from '../src/domain/panels.js';

const GENRES = [
  { id: 'none', label: '未設定' },
  { id: 'meal', label: '食事関連' }
];

const makeSheet = (genre, overrides = {}) => {
  const panels = buildDefaultPanels();
  Object.entries(overrides).forEach(([index, patch]) => {
    panels[Number(index)] = { ...panels[Number(index)], ...patch };
  });
  return { id: `sheet-${genre}`, genre, panels };
};

test('buildPageCsvContent starts with BOM and the fixed header row', () => {
  const content = buildPageCsvContent({ sheets: [], genres: GENRES });
  assert.equal(content.charCodeAt(0), 0xfeff);
  assert.equal(content.slice(1), PAGE_CSV_HEADERS.join(','));
  assert.deepEqual(PAGE_CSV_HEADERS, [
    'ジャンル', '介援隊コード', 'ページ数', 'X_POS', 'Y_POS', 'コマ番号', 'コマ数', 'コマID', 'テキスト情報'
  ]);
  assert.equal(PAGE_CSV_HEADERS.includes('追番'), false);
  assert.equal(PAGE_CSV_HEADERS.includes('座標'), false);
});

test('buildPageCsvRowsForSheet numbers visible panels and skips hidden ones', () => {
  const sheet = makeSheet('meal', {
    0: { code: 'A1234', rowSpan: 1, colSpan: 2, sizeType: null }, // sizeType 未設定時は rowSpan/colSpan から算出
    1: { hidden: true },
    2: { label: 'タイトル' },
    5: { text: 'テキスト, 付き', isText: true }
  });

  const rows = buildPageCsvRowsForSheet(sheet, { pageNum: 3, genreLabel: '食事関連' });
  assert.equal(rows.length, 15); // 16 コマ - hidden 1

  // 先頭コマ: X_POS=1 / Y_POS=1 / コマ番号1
  assert.equal(rows[0], '食事関連,A1234,3,1,1,1,1/8 横（2コマ）,,');
  // 特殊ダミー(タイトル)もコマ番号を持ち、コードは「ダミーコマ」
  assert.equal(rows[1], '食事関連,ダミーコマ,3,3,1,2,1/16（1コマ）,,');
  // テキスト情報はカンマを含むのでダブルクォートで囲まれる
  assert.equal(rows[4], '食事関連,,3,2,2,5,1/16（1コマ）,,"テキスト, 付き"');
});

test('buildPageCsvContent resolves genre labels per sheet and falls back to 未設定', () => {
  const content = buildPageCsvContent({
    sheets: [makeSheet('meal'), makeSheet('unknown-genre')],
    genres: GENRES
  });
  const lines = content.split('\n');
  assert.equal(lines.length, 1 + 32);
  assert.ok(lines[1].startsWith('食事関連,,1,'));
  assert.ok(lines[17].startsWith('未設定,,2,'));
});

test('buildExcludedItemsCsvContent formats createdAt from seconds and falls back to now', () => {
  const now = new Date('2026-08-31T00:00:00Z');
  const content = buildExcludedItemsCsvContent([
    { code: 'B0001', originalName: 'b.jpg', label: 'ラベル', createdAt: { seconds: 1_700_000_000 } },
    { code: 'B0002' }
  ], { now });
  const lines = content.split('\n');
  assert.equal(lines[0].slice(1), '介援隊コード,画像名,ラベル,登録日時');
  assert.equal(lines[1], `B0001,b.jpg,ラベル,${new Date(1_700_000_000 * 1000).toLocaleString()}`);
  assert.equal(lines[2], `B0002,,,${now.toLocaleString()}`);
});

test('buildDatedCsvFilename appends ISO date', () => {
  assert.equal(buildDatedCsvFilename('daiwari_export', new Date('2026-08-31T12:34:56Z')), 'daiwari_export_2026-08-31.csv');
});
