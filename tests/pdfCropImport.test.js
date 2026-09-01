import test from 'node:test';
import assert from 'node:assert/strict';

import {
  DEFAULT_PDF_GRID_BOUNDS,
  getPdfCropRect,
  inferCatalogStartPage,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode,
  parsePdfCropCsv,
  pdfTextItemsContainCode,
  updatePdfCropRowSize
} from '../src/domain/pdfCropImport.js';

const HEADER = 'ジャンル,ページ数,追番,コマ番号,介援隊コード,コマ数,,テキスト情報,座標,コマID,X_POS,Y_POS';

test('normalizePdfCropCode removes catalog prefixes and proof suffixes', () => {
  assert.equal(normalizePdfCropCode('261-E1957'), 'E1957');
  assert.equal(normalizePdfCropCode('010-E1894_'), 'E1894');
  assert.equal(normalizePdfCropCode('ａ１２３４'), 'A1234');
  assert.equal(normalizePdfCropCode('ダミーコマ'), '');
});

test('inferCatalogStartPage reads the P number from a proof PDF filename', () => {
  assert.equal(inferCatalogStartPage('P010_介援隊vol.26.pdf'), 10);
  assert.equal(inferCatalogStartPage('catalog.pdf'), 1);
});

test('parsePdfCropCsv reads code, size, page and explicit grid coordinates', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,261-E1957,1/8 横（2コマ）,,,X1Y1,PID-1,1,1',
    '食事関連,10,2,2,E1894,1/8 横（2コマ）,,,X3Y1,PID-2,3,1',
    '食事関連,10,3,3,ダミーコマ,1/16（1コマ）,タイトル,,X1Y2,PID-3,1,2',
    '食事関連,10,4,4,E9999,1/16（1コマ）,,説明文,X2Y2,PID-4,2,2'
  ].join('\n');

  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(issues.length, 0);
  assert.equal(rows.length, 2, 'dummy and text rows are not image crop targets');
  assert.deepEqual(rows[0], {
    id: '10-E1957-2',
    csvRow: 2,
    pageNumber: 10,
    order: 1,
    frameNumber: 1,
    rawCode: '261-E1957',
    code: 'E1957',
    filename: 'E1957.jpg',
    sizeType: '1/8 横（2コマ）',
    rowSpan: 1,
    colSpan: 2,
    xPos: 1,
    yPos: 1,
    positionSource: 'csv',
    layoutStatus: 'ready'
  });
});

test('parsePdfCropCsv estimates missing positions in frame order', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,E1957,1/8 横（2コマ）,,,,,,',
    '食事関連,10,2,2,E1894,1/8 横（2コマ）,,,,,,'
  ].join('\n');
  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(issues.length, 0);
  assert.deepEqual(rows.map(({ xPos, yPos, positionSource }) => ({ xPos, yPos, positionSource })), [
    { xPos: 1, yPos: 1, positionSource: 'estimated' },
    { xPos: 3, yPos: 1, positionSource: 'estimated' }
  ]);
});

test('parsePdfCropCsv reports overlapping explicit positions', () => {
  const content = [
    HEADER,
    '食事関連,10,1,1,E1957,1/8 横（2コマ）,,,X1Y1,,,',
    '食事関連,10,2,2,E1894,1/16（1コマ）,,,X2Y1,,,'
  ].join('\n');
  const { rows, issues } = parsePdfCropCsv(content);
  assert.equal(rows[1].layoutStatus, 'conflict');
  assert.equal(issues[0].type, 'layout-conflict');
});

test('getPdfCropRect converts 4x4 coordinates and spans to normalized bounds', () => {
  const row = { xPos: 3, yPos: 2, rowSpan: 2, colSpan: 2, layoutStatus: 'ready' };
  const rect = getPdfCropRect(row, { left: 10, top: 6, right: 10, bottom: 10 });
  assert.deepEqual(rect, { x: 0.5, y: 0.27, width: 0.4, height: 0.42 });
  assert.equal(isPdfCropRowInsideGrid(row), true);
  assert.equal(isPdfCropRowInsideGrid({ ...row, xPos: 4 }), false);
});

test('row size edits update spans and text verification is limited to its crop', () => {
  const base = { code: 'E1957', xPos: 1, yPos: 1, rowSpan: 1, colSpan: 1, layoutStatus: 'ready' };
  const resized = updatePdfCropRowSize(base, '1/8 横（2コマ）');
  assert.equal(resized.colSpan, 2);
  assert.equal(pdfTextItemsContainCode([
    { text: '261-', x: 0.11, y: 0.2 },
    { text: 'E1957', x: 0.2, y: 0.2 },
    { text: 'E1957', x: 0.8, y: 0.8 }
  ], resized, DEFAULT_PDF_GRID_BOUNDS), true);
  assert.equal(pdfTextItemsContainCode([{ text: 'E1957', x: 0.8, y: 0.8 }], resized), false);
});
