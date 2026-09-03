import test from 'node:test';
import assert from 'node:assert/strict';

import {
  CSV_BOM,
  buildBomCsvContent,
  escapeCsvCell,
  parseCSVLine
} from '../src/lib/csv.js';

test('CSV cell escaping preserves empty and plain values and quotes CSV-sensitive content', () => {
  assert.equal(escapeCsvCell(null), '');
  assert.equal(escapeCsvCell(0), '0');
  assert.equal(escapeCsvCell('plain'), 'plain');
  assert.equal(escapeCsvCell('with,comma'), '"with,comma"');
  assert.equal(escapeCsvCell('say "hello"'), '"say ""hello"""');
  assert.equal(escapeCsvCell('line 1\r\nline 2'), '"line 1\r\nline 2"');
});

test('CSV cell escaping can retain the legacy page CSV handling of a standalone carriage return', () => {
  assert.equal(escapeCsvCell('line 1\rline 2', { quoteCarriageReturn: false }), 'line 1\rline 2');
});

test('BOM CSV builder preserves BOM, LF separators, empty rows, and optional header escaping', () => {
  assert.equal(
    buildBomCsvContent(['name', 'value,unit'], ['item,1', '']),
    `${CSV_BOM}name,value,unit\nitem,1\n`
  );
  assert.equal(
    buildBomCsvContent(['name', 'value,unit'], ['item,1'], { escapeHeaders: true }),
    `${CSV_BOM}name,"value,unit"\nitem,1`
  );
});

test('CSV parser handles commas, escaped quotes, empty values, and trimming', () => {
  assert.deepEqual(
    parseCSVLine(' first ,"second, value","escaped ""quote""",, last '),
    ['first', 'second, value', 'escaped "quote"', '', 'last']
  );
});

test('CSV parser preserves line content inside a quoted field', () => {
  assert.deepEqual(parseCSVLine('"line 1\nline 2",value'), ['line 1\nline 2', 'value']);
});
