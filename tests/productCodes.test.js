import test from 'node:test';
import assert from 'node:assert/strict';

import { extractProductCodes, normalizeCode } from '../src/domain/productCodes.js';

test('product codes normalize full-width characters, separators, and case', () => {
  assert.equal(normalizeCode(' ａｂ-１２ ３４ '), 'AB1234');
  assert.equal(normalizeCode('a-1234'), 'A1234');
  assert.equal(normalizeCode('Ａ－１２３４'), 'A－1234');
  assert.equal(normalizeCode(''), '');
  assert.equal(normalizeCode(null), '');
});

test('extractProductCodes collects every code written in a panel', () => {
  // 区切り文字ちがい・カタログ前置き・全角のいずれでも拾える
  assert.deepEqual(extractProductCodes('E1760 / E1761'), ['E1760', 'E1761']);
  assert.deepEqual(extractProductCodes('261-E0340'), ['E0340']);
  assert.deepEqual(extractProductCodes('ｅ１７６０・E1761'), ['E1760', 'E1761']);
  // コード欄とテキスト欄をまとめて渡せて、重複は 1 回だけ
  assert.deepEqual(extractProductCodes('E1760', 'E1760とE1761のセット'), ['E1760', 'E1761']);
  // コードの形をしていない文字列は無視する
  assert.deepEqual(extractProductCodes('タイトル', '500ml', null, ''), []);
  assert.deepEqual(extractProductCodes(), []);
});
