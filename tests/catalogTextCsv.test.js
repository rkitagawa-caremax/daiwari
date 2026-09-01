import test from 'node:test';
import assert from 'node:assert/strict';

import {
  CATALOG_TEXT_CSV_HEADERS,
  buildCatalogTextCsvContent,
  countCatalogTextExportImages,
  resolveCatalogTextExportData
} from '../src/domain/catalogTextCsv.js';
import { parseCSVLine } from '../src/lib/csv.js';

const valueByHeader = (headers, values, header) => values[headers.indexOf(header)];

test('catalog text CSV is BOM-prefixed, excludes ordinary uploads and sorts PDF crops by page', () => {
  const images = [
    { id: 'manual', name: 'manual.jpg', data: 'data:manual' },
    { id: 'page-12', name: 'E1955.jpg', code: 'E1955', sourcePdfName: 'P012.pdf', sourcePage: 12, sourceText: '261-E1955 ¥5,832' },
    { id: 'page-10', name: 'E1894.jpg', code: 'E1894', sourcePdfName: 'P010.pdf', sourcePage: 10, sourceText: '261-E1894 ¥2,139' }
  ];
  const content = buildCatalogTextCsvContent(images);
  const lines = content.split('\n');

  assert.equal(content.charCodeAt(0), 0xfeff);
  assert.deepEqual(parseCSVLine(lines[0].slice(1)), CATALOG_TEXT_CSV_HEADERS);
  assert.equal(countCatalogTextExportImages(images), 2);
  assert.ok(lines[1].includes('E1894'));
  assert.ok(lines[2].includes('E1955'));
});

test('legacy source text is reanalyzed into item number, prices, markers and specifications', () => {
  const image = {
    id: 'legacy-pdf',
    name: 'E0503.jpg',
    code: 'E0503',
    sourcePdfName: 'P036.pdf',
    sourcePage: 36,
    pdfPageNumber: 1,
    sizeType: '1/16（1コマ）',
    sourceText: '使い捨て防水エプロン 261-E0503 KN-932 ¥1,028 (税抜¥935) ●材質/パルプ、ポリエチレン ●成分/パルプ (D) 在庫商品'
  };
  const content = buildCatalogTextCsvContent([image]);
  const [headerLine, rowLine] = content.split('\n');
  const headers = parseCSVLine(headerLine.slice(1));
  const row = parseCSVLine(rowLine);

  assert.equal(valueByHeader(headers, row, '品番'), 'KN-932');
  assert.equal(valueByHeader(headers, row, '税込価格'), '1028');
  assert.equal(valueByHeader(headers, row, '税抜価格'), '935');
  assert.equal(valueByHeader(headers, row, '販売区分'), '在庫');
  assert.equal(valueByHeader(headers, row, '取扱記号'), '(D)');
  assert.equal(valueByHeader(headers, row, 'デモ機'), 'あり');
  assert.match(valueByHeader(headers, row, '仕様'), /●材質\/パルプ、ポリエチレン/);
  assert.equal(valueByHeader(headers, row, '成分・原材料・栄養'), '●成分/パルプ (D)');
});

test('saved structured values take priority and CSV-sensitive source text is escaped', () => {
  const image = {
    id: 'current-pdf',
    name: 'E1919.jpg',
    code: 'E1919',
    productName: 'ケアハート「お口,潤う」泡スプレー',
    productNameSource: 'csv',
    sourcePdfName: 'P030.pdf',
    sourceText: '261-E1919 OLD-001 ¥877, "原文"',
    priceIncludingTax: 877,
    priceExtractionConfidence: 'high',
    catalogTextData: {
      version: 1,
      itemNumber: 'SAVED-001',
      itemNumberCandidates: ['SAVED-001'],
      catchCopy: '保存済みコピー。',
      catchCopyCandidates: ['保存済みコピー。'],
      availability: 'direct',
      availabilityLabels: ['直送'],
      handlingMarkers: ['(S)'],
      hasDemoMarker: false,
      specifications: ['●成分/水、グリセリン'],
      compositionDetails: ['●成分/水、グリセリン'],
      materialDetails: []
    }
  };
  const resolved = resolveCatalogTextExportData(image);
  const content = buildCatalogTextCsvContent([image]);
  const [headerLine, rowLine] = content.split('\n');
  const headers = parseCSVLine(headerLine.slice(1));
  const row = parseCSVLine(rowLine);

  assert.equal(resolved.details.itemNumber, 'SAVED-001');
  assert.equal(valueByHeader(headers, row, '商品名'), image.productName);
  assert.equal(valueByHeader(headers, row, '品番'), 'SAVED-001');
  assert.equal(valueByHeader(headers, row, '販売区分'), '直送');
  assert.equal(valueByHeader(headers, row, '元テキスト'), image.sourceText);
});
