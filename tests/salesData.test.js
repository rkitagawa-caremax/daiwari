import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildMonthlySalesSeries,
  mergeSerializedSalesChunks,
  parseSalesCsvContent,
  resolveSalesMonthColumns,
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

test('sales CSV keeps fiscal-year monthly values without changing the total', () => {
  const yearHeader = Array(18).fill('');
  const monthHeader = Array(18).fill('');
  yearHeader[4] = '2025年';
  ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
    .forEach((label, index) => { monthHeader[index + 4] = label; });
  const dataRow = Array(18).fill('');
  dataRow[1] = '商品A';
  dataRow[2] = '規格A';
  dataRow[3] = 'E001';
  Array.from({ length: 12 }, (_, index) => index + 1)
    .forEach((count, index) => { dataRow[index + 4] = String(count); });
  dataRow[17] = '78';

  const parsed = parseSalesCsvContent([
    yearHeader.join(','),
    monthHeader.join(','),
    dataRow.join(',')
  ].join('\n'));

  assert.equal(parsed.E001[0].count, 78);
  assert.equal(parsed.E001[0].monthlySales.length, 12);
  assert.deepEqual(parsed.E001[0].monthlySales.slice(0, 2), [1, 2]);
  assert.equal(parsed.E001[0].monthlySales.at(-1), 12);
  assert.deepEqual(parsed.E001[0].monthlyLabels, [
    '25/4', '25/5', '25/6', '25/7', '25/8', '25/9',
    '25/10', '25/11', '25/12', '26/1', '26/2', '26/3'
  ]);
});

test('monthly sales series aggregates product variations in source column order', () => {
  const items = [
    { monthlySales: [{ key: '2026-01', label: '26/1', year: 2026, month: 1, count: 10 }, { key: '2026-02', label: '26/2', year: 2026, month: 2, count: 4 }] },
    { monthlySales: [{ key: '2026-01', label: '26/1', year: 2026, month: 1, count: 5 }, { key: '2026-02', label: '26/2', year: 2026, month: 2, count: 8 }] }
  ];
  assert.deepEqual(buildMonthlySalesSeries(items), [
    { key: '2026-01', label: '26/1', year: 2026, month: 1, count: 15 },
    { key: '2026-02', label: '26/2', year: 2026, month: 2, count: 12 }
  ]);
});

test('monthly sales series aggregates compact persisted values using one shared label list', () => {
  assert.deepEqual(buildMonthlySalesSeries([
    { monthlyLabels: ['25/4', '25/5'], monthlySales: [10, 4] },
    { monthlySales: [5, 8] }
  ]), [
    { key: 'compact-month-0', label: '25/4', year: null, month: null, count: 15 },
    { key: 'compact-month-1', label: '25/5', year: null, month: null, count: 12 }
  ]);
});

test('month detection ignores total columns', () => {
  assert.deepEqual(resolveSalesMonthColumns([
    ['', '', '', '', '2026年1月', '2026年2月', '12ヶ月合計']
  ]).map((column) => column.columnIndex), [4, 5]);
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

test('splitSalesDataIntoChunks also limits UTF-8 serialized byte size', () => {
  const salesData = {
    A: [{ name: 'あ'.repeat(40), count: 1 }],
    B: [{ name: 'い'.repeat(40), count: 2 }]
  };
  const chunks = splitSalesDataIntoChunks(salesData, 1000, 180);
  assert.deepEqual(chunks.map((chunk) => Object.keys(chunk)), [['A'], ['B']]);
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
