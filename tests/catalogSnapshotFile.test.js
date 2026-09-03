import test from 'node:test';
import assert from 'node:assert/strict';

import { parseBestWorkbookSheet } from '../src/lib/catalogSnapshotFile.js';

test('workbook import chooses the sheet with the most recognized catalog fields', () => {
  const result = parseBestWorkbookSheet([
    { sheet: '説明', data: [['注意事項'], ['介援隊コード'], ['E001']] },
    { sheet: '価格改定', data: [['介援隊コード', '税込価格', '販売状態'], ['E001', 1540, '在庫限り']] }
  ]);
  assert.equal(result.sheetName, '価格改定');
  assert.deepEqual(result.recognizedFields, ['priceIncludingTax', 'lifecycleStatus']);
  assert.equal(result.items[0].lifecycleStatus, '在庫限り');
});
