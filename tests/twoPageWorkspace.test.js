import assert from 'node:assert/strict';
import test from 'node:test';

import {
  getAdjacentPageOptions,
  getPageNavigationSelection,
  getTwoPageDisplaySheets
} from '../src/domain/twoPageWorkspace.js';

const sheets = [
  { id: 'page-a' },
  { id: 'page-b' },
  { id: 'page-c' }
];

test('前後のページだけをページ番号付きの候補にする', () => {
  assert.deepEqual(getAdjacentPageOptions(sheets, 'page-b'), [
    { id: 'page-a', pageNumber: 1, direction: 'prev', directionLabel: '前のページ' },
    { id: 'page-c', pageNumber: 3, direction: 'next', directionLabel: '次のページ' }
  ]);
});

test('先頭と末尾では存在する側だけを候補にする', () => {
  assert.deepEqual(getAdjacentPageOptions(sheets, 'page-a').map((option) => option.id), ['page-b']);
  assert.deepEqual(getAdjacentPageOptions(sheets, 'page-c').map((option) => option.id), ['page-b']);
});

test('現在ページを先頭にして選択ページを追加表示する', () => {
  assert.deepEqual(
    getTwoPageDisplaySheets(sheets, 'page-b', 'page-a').map((sheet) => sheet.id),
    ['page-b', 'page-a']
  );
});

test('重複・削除済みの追加ページは表示しない', () => {
  assert.deepEqual(
    getTwoPageDisplaySheets(sheets, 'page-b', 'page-b').map((sheet) => sheet.id),
    ['page-b']
  );
  assert.deepEqual(
    getTwoPageDisplaySheets(sheets, 'page-b', 'missing').map((sheet) => sheet.id),
    ['page-b']
  );
});

test('2P表示中のページ移動では隣接した2ページを保つ', () => {
  assert.deepEqual(
    getPageNavigationSelection(sheets, 'page-a', 'page-b', 'next'),
    { activeSheetId: 'page-b', secondarySheetId: 'page-c' }
  );
  assert.deepEqual(
    getPageNavigationSelection(sheets, 'page-c', 'page-b', 'prev'),
    { activeSheetId: 'page-b', secondarySheetId: 'page-a' }
  );
});

test('端では追加ページを反対側へ移して2P表示を維持する', () => {
  assert.deepEqual(
    getPageNavigationSelection(sheets, 'page-b', 'page-c', 'next'),
    { activeSheetId: 'page-c', secondarySheetId: 'page-b' }
  );
  assert.equal(getPageNavigationSelection(sheets, 'page-a', 'page-b', 'prev'), null);
});
