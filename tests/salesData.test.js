import test from 'node:test';
import assert from 'node:assert/strict';

import {
  mergeSerializedSalesChunks,
  parseSalesCsvContent,
  splitSalesDataIntoChunks
} from '../src/domain/salesData.js';

const makeSalesRow = ({ name = '', spec = '', code = '', count = '' } = {}) => {
  const columns = Array(18).fill('');
  columns[1] = name;
  columns[2] = spec;
  columns[3] = code;
  columns[17] = count;
  return columns.join(',');
};

test('parseSalesCsvContent skips two header rows and normalizes product codes', () => {
  const csv = [
    'metadata',
    'header',
    makeSalesRow({ name: '商品A', spec: '10個入', code: 'ＡＢ-１２３', count: '15' }),
    makeSalesRow({ name: '商品B', spec: '20個入', code: 'ab 123', count: '5' })
  ].join('\r\n');

  assert.deepEqual(parseSalesCsvContent(csv), {
    AB123: [
      { name: '商品A', spec: '10個入', count: 15 },
      { name: '商品B', spec: '20個入', count: 5 }
    ]
  });
});

test('parseSalesCsvContent handles quoted commas and ignores incomplete rows', () => {
  const quotedRow = makeSalesRow({
    name: '"商品, A"',
    spec: '規格',
    code: 'C001',
    count: '"1,234"'
  });
  const csv = ['metadata', 'header', quotedRow, 'short,row', ''].join('\n');

  assert.deepEqual(parseSalesCsvContent(csv), {
    C001: [{ name: '商品, A', spec: '規格', count: 1234 }]
  });
});

test('splitSalesDataIntoChunks keeps entry order and chunk boundaries', () => {
  const salesData = {
    A: [{ count: 1 }],
    B: [{ count: 2 }],
    C: [{ count: 3 }]
  };

  assert.deepEqual(splitSalesDataIntoChunks(salesData, 2), [
    { A: [{ count: 1 }], B: [{ count: 2 }] },
    { C: [{ count: 3 }] }
  ]);
});

test('mergeSerializedSalesChunks merges valid chunks and reports invalid chunks', () => {
  const parseErrors = [];
  const merged = mergeSerializedSalesChunks([
    JSON.stringify({ A: [{ count: 1 }] }),
    '{invalid json',
    JSON.stringify({ B: [{ count: 2 }] })
  ], {
    onParseError: (_error, index) => parseErrors.push(index)
  });

  assert.deepEqual(merged, {
    A: [{ count: 1 }],
    B: [{ count: 2 }]
  });
  assert.deepEqual(parseErrors, [1]);
});
