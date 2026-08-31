import test from 'node:test';
import assert from 'node:assert/strict';

import {
  IMAGE_MATCH_THRESHOLD,
  buildImportedSheets,
  buildPageCsvImportReport,
  buildSearchableImages,
  findBestImageMatch,
  parsePageCsvRows
} from '../src/domain/pageCsvImport.js';
import { parseCSVLine } from '../src/lib/csv.js';

const GENRES = [
  { id: 'none', label: '未設定' },
  { id: 'meal', label: '食事関連' },
  { id: 'bath', label: '入浴関連' }
];

const IMAGES = [
  { id: 'img-a1234', name: 'A1234.jpg', data: 'data:...' },
  { id: 'img-b0001', name: 'b-0001_photo.png', data: 'data:...' },
  { id: 'img-1234', name: '1234.jpg', data: 'data:...' }
];

const parseRows = (rows, extra = {}) => parsePageCsvRows(rows, { parseLine: parseCSVLine, images: IMAGES, genres: GENRES, ...extra });

const HEADER = 'ジャンル,ページ数,追番,コマ番号,介援隊コード,コマ数,,テキスト情報,座標,コマID';

test('buildSearchableImages keeps only id and name', () => {
  assert.deepEqual(buildSearchableImages(IMAGES)[0], { id: 'img-a1234', name: 'A1234.jpg' });
  assert.deepEqual(buildSearchableImages([{ id: 'x' }]), [{ id: 'x', name: '' }]);
});

test('findBestImageMatch scores exact stem match as 100 and ignores extension/case/full-width', () => {
  const images = buildSearchableImages(IMAGES);
  assert.deepEqual(findBestImageMatch('A1234', images), { bestMatchImg: images[0], bestScore: 100 });
  assert.equal(findBestImageMatch('ａ１２３４', images).bestScore, 100);
});

test('findBestImageMatch: numeric token match scores 80 (single token) and stays under the threshold for symbol-only matches', () => {
  const images = buildSearchableImages([{ id: 'n', name: 'photo_0777.jpg' }, { id: 'm', name: 'ab-12.jpg' }]);
  assert.deepEqual(findBestImageMatch('777', images), { bestMatchImg: images[0], bestScore: 80 });
  const symbolOnly = findBestImageMatch('AB12', images);
  assert.equal(symbolOnly.bestScore, 50);
  assert.ok(symbolOnly.bestScore < IMAGE_MATCH_THRESHOLD);
  assert.deepEqual(findBestImageMatch('ZZZ', images), { bestMatchImg: null, bestScore: 0 });
});

test('parsePageCsvRows resolves text / dummy / matched code / unmatched code rows', async () => {
  const rows = [
    HEADER,
    '食事関連,1,1,1,,1/16（1コマ）,,自由テキスト,X1Y1,',
    '食事関連,1,2,2,ダミーコマ,1/16（1コマ）,タイトル,,X2Y1,',
    '食事関連,1,3,3,A1234,1/8 横（2コマ）,,,X3Y1,PID-3',
    '食事関連,1,4,4,Z9999,1/16（1コマ）,,,X1Y2,',
    '食事関連,1,5,5,自由ラベル,1/16（1コマ）,,,X2Y2,',
    '',
    '入浴関連,,1,1,A1234,1/16（1コマ）,,,X1Y1,',   // ページ番号なし → skip
    '入浴関連,3,,,A1234,1/16（1コマ）,,,X1Y1,'     // 追番・コマ番号なし → skip
  ];

  const { sheetUpdates, maxPageIndex } = await parseRows(rows);
  assert.equal(maxPageIndex, 0);
  assert.deepEqual(Object.keys(sheetUpdates), ['0']);

  const page = sheetUpdates[0];
  assert.equal(page.genre, 'meal');
  assert.equal(page.contentItems.length, 5);

  const [textItem, dummyItem, codeItem, unmatchedItem, labelItem] = page.contentItems;
  assert.deepEqual(textItem, {
    isFixed: true, frameNo: 1, order: 1,
    data: { code: null, image: null, imageId: null, label: 'テキスト', sizeType: '1/16（1コマ）', text: '自由テキスト', isText: true, panelId: null }
  });
  assert.equal(dummyItem.data.label, 'タイトル');
  assert.equal(dummyItem.data.code, null);
  assert.equal(codeItem.data.code, 'A1234');
  assert.equal(codeItem.data.imageId, 'img-a1234');
  assert.equal(codeItem.data.panelId, 'PID-3');
  assert.equal(unmatchedItem.data.code, 'Z9999');
  assert.equal(unmatchedItem.data.imageId, null);
  assert.equal(labelItem.data.label, '自由ラベル');
  assert.equal(labelItem.data.code, null);

  assert.deepEqual(page.matchDetails, [{ code: 'A1234', imageName: 'A1234.jpg', score: 100, csvRow: 4 }]);
  assert.deepEqual(page.unmatchedCodes, [{ code: 'Z9999', csvRow: 5, bestScore: 0, bestMatch: 'なし' }]);
});

test('parsePageCsvRows falls back between 追番 and コマ番号 and reports progress every 50 rows', async () => {
  const rows = [HEADER];
  for (let i = 1; i <= 120; i++) {
    // 追番のみ (コマ番号なし) → isFixed true (frameNum は追番から補完)
    rows.push(`未設定,${i},2,,ダミーコマ,1/16（1コマ）,,,,`);
  }
  const progress = [];
  const { sheetUpdates, maxPageIndex } = await parseRows(rows, { onProgress: (i, total) => { progress.push([i, total]); } });
  assert.equal(maxPageIndex, 119);
  assert.deepEqual(progress, [[50, 121], [100, 121]]);
  assert.deepEqual(sheetUpdates[0].contentItems[0], {
    isFixed: true, frameNo: 2, order: 2,
    data: { code: null, image: null, imageId: null, label: '新規商品未確定', sizeType: '1/16（1コマ）', text: '', isText: false, panelId: null }
  });
});

test('buildImportedSheets pads missing pages, places panels in order, and leaves existing sheets untouched', async () => {
  const existing = [{ id: 'sheet-1', genre: 'none', panels: [] }];
  const sheetUpdates = {
    1: {
      genre: 'bath',
      contentItems: [
        { isFixed: true, frameNo: 2, order: 2, data: { code: null, image: null, imageId: null, label: '埋草', sizeType: '1/16（1コマ）', text: '', isText: false, panelId: null } },
        { isFixed: true, frameNo: 1, order: 1, data: { code: 'A1234', image: null, imageId: 'img-a1234', label: null, sizeType: '1/8 横（2コマ）', text: '', isText: false, panelId: null } }
      ],
      matchDetails: [{ code: 'A1234', imageName: 'A1234.jpg', score: 100, csvRow: 3 }, { code: 'A1234', imageName: 'A1234.jpg', score: 100, csvRow: 4 }],
      unmatchedCodes: [{ code: 'Z9999', csvRow: 5, bestScore: 0, bestMatch: 'なし' }]
    }
  };
  let counter = 0;
  const progress = [];
  const { localSheets, importSummary } = await buildImportedSheets({
    sheets: existing,
    sheetUpdates,
    finalPageCount: 2,
    generateId: () => `local_${++counter}`,
    now: () => 1_000_000,
    onProgress: (i, total) => { progress.push([i, total]); }
  });

  assert.equal(existing.length, 1, 'input array is not mutated');
  assert.equal(localSheets.length, 2);
  assert.equal(localSheets[0], existing[0]);
  assert.equal(localSheets[1].id, 'local_1');
  assert.equal(localSheets[1].genre, 'bath');
  assert.deepEqual(localSheets[1].createdAt, { seconds: 1000 });
  assert.deepEqual(progress, [[1, 2]]);

  const panels = localSheets[1].panels;
  // コマ1 (1/8 横) が X1Y1 を起点に 2 マス占有 → コマ2 (埋草) は次の空き X3Y1 (index 2)
  assert.equal(panels[0].code, 'A1234');
  assert.equal(panels[0].colSpan, 2);
  assert.equal(panels[1].hidden, true);
  assert.equal(panels[2].label, '埋草');
  assert.equal(panels[3].hidden, false);

  assert.equal(importSummary.total, 2);
  assert.equal(importSummary.autoSuccess, 2);
  assert.equal(importSummary.autoFailed, 0);
  assert.equal(importSummary.matchedImages.length, 2);
  assert.deepEqual(importSummary.imageUsageCount, { 'A1234.jpg': 2 });
  assert.deepEqual(importSummary.notMatchedCodes, [{ page: 2, code: 'Z9999', csvRow: 5, bestScore: 0, bestMatch: 'なし' }]);
});

test('buildImportedSheets records items that do not fit', async () => {
  const contentItems = Array.from({ length: 17 }, (_, index) => ({
    isFixed: true, frameNo: index + 1, order: index + 1,
    data: { code: null, image: null, imageId: null, label: 'X', sizeType: '1/16（1コマ）', text: '', isText: false, panelId: null }
  }));
  const { importSummary } = await buildImportedSheets({
    sheets: [],
    sheetUpdates: { 0: { genre: null, contentItems } },
    finalPageCount: 1,
    generateId: () => 'local_x'
  });
  assert.equal(importSummary.autoSuccess, 16);
  assert.equal(importSummary.autoFailed, 1);
  assert.equal(importSummary.details.length, 1);
  assert.match(importSummary.details[0], /ページ 1: 「X」/);
});

test('buildPageCsvImportReport lists counts, unmatched codes, duplicates, and unplaced items', () => {
  const report = buildPageCsvImportReport({
    total: 3, fixedSuccess: 0, fixedFailed: 0, autoSuccess: 2, autoFailed: 1,
    details: ['・ページ 1: 「X」（1/16（1コマ）） を配置できませんでした。'],
    matchedImages: [{}, {}],
    notMatchedCodes: [{ page: 1, csvRow: 5, code: 'Z9999', bestMatch: 'なし', bestScore: 0 }],
    imageUsageCount: { 'A1234.jpg': 2, 'B0001.jpg': 1 }
  });
  const lines = report.split('\n');
  assert.equal(lines[0], '取り込みが完了しました。(全3件)');
  assert.ok(lines.includes('・スペース不足で配置失敗: 1件'));
  assert.ok(lines.includes('・マッチング成功: 2件'));
  assert.ok(lines.includes('・P1 行5: Z9999 (ベストマッチ: なし, スコア: 0)'));
  assert.ok(lines.includes('・A1234.jpg (2回)'));
  assert.ok(!report.includes('B0001.jpg'));
  assert.ok(report.includes('【未配置の項目】'));
});
