import test from 'node:test';
import assert from 'node:assert/strict';

import { parseCSVLine } from '../src/lib/csv.js';

test('CSV parser handles commas, escaped quotes, empty values, and trimming', () => {
  assert.deepEqual(
    parseCSVLine(' first ,"second, value","escaped ""quote""",, last '),
    ['first', 'second, value', 'escaped "quote"', '', 'last']
  );
});

test('CSV parser preserves line content inside a quoted field', () => {
  assert.deepEqual(parseCSVLine('"line 1\nline 2",value'), ['line 1\nline 2', 'value']);
});
