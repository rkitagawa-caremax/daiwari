import test from 'node:test';
import assert from 'node:assert/strict';

import { normalizeCode } from '../src/domain/productCodes.js';

test('product codes normalize full-width characters, separators, and case', () => {
  assert.equal(normalizeCode(' ａｂ-１２ ３４ '), 'AB1234');
  assert.equal(normalizeCode('a-1234'), 'A1234');
  assert.equal(normalizeCode('Ａ－１２３４'), 'A－1234');
  assert.equal(normalizeCode(''), '');
  assert.equal(normalizeCode(null), '');
});
