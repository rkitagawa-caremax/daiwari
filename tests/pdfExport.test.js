import test from 'node:test';
import assert from 'node:assert/strict';

import { buildPdfExportPlan, sanitizePdfFilenamePart } from '../src/domain/pdfExport.js';

const genres = [
  { id: 'meal', label: '食事関連' },
  { id: 'bath', label: '入浴関連' }
];

const sheets = [
  { id: 'sheet-1', genre: 'meal' },
  { id: 'sheet-2', genre: 'bath' },
  { id: 'sheet-3', genre: 'meal' }
];

test('PDF export plan keeps selected pages in actual page order', () => {
  const plan = buildPdfExportPlan({
    sheets,
    selectedSheetIds: new Set(['sheet-3', 'sheet-1']),
    genres
  });

  assert.deepEqual(plan.pages.map((page) => page.sheet.id), ['sheet-1', 'sheet-3']);
  assert.deepEqual(plan.pages.map((page) => page.pageNumber), [1, 3]);
  assert.equal(plan.filename, '食事関連.pdf');
});

test('single-page PDF filename contains page number and genre', () => {
  const plan = buildPdfExportPlan({ sheets, selectedSheetIds: ['sheet-2'], genres });

  assert.equal(plan.filename, 'Page2_入浴関連.pdf');
});

test('mixed-genre PDF filename contains genres only and joins into one PDF name', () => {
  const plan = buildPdfExportPlan({
    sheets,
    selectedSheetIds: ['sheet-1', 'sheet-2', 'sheet-3'],
    genres
  });

  assert.equal(plan.filename, '食事関連・入浴関連.pdf');
});

test('PDF filename sanitization removes filesystem-reserved characters', () => {
  assert.equal(sanitizePdfFilenamePart(' 医療/施設:*? '), '医療_施設___');
  assert.deepEqual(buildPdfExportPlan({ sheets, selectedSheetIds: [], genres }), { pages: [], filename: '' });
});
