import assert from 'node:assert/strict';
import test from 'node:test';

import {
  FREE_LABEL_MAX_CHARACTERS,
  FREE_LABEL_MAX_COLUMNS,
  FREE_LABEL_MAX_ROWS,
  constrainFreeLabelText,
  formatFreeLabelText,
  getFreeLabelTextLayout
} from '../src/domain/freeLabels.js';

test('free label text is limited to 80 visible characters', () => {
  const constrained = constrainFreeLabelText('あ'.repeat(90));

  assert.equal(Array.from(constrained).length, FREE_LABEL_MAX_CHARACTERS);
  assert.equal(constrained, 'あ'.repeat(80));
});

test('free label layout wraps at 8 characters and expands to 10 rows', () => {
  const layout = getFreeLabelTextLayout('あ'.repeat(80));

  assert.deepEqual(layout, {
    characterCount: FREE_LABEL_MAX_CHARACTERS,
    columns: FREE_LABEL_MAX_COLUMNS,
    rows: FREE_LABEL_MAX_ROWS
  });
});

test('free label layout respects explicit line breaks', () => {
  assert.deepEqual(getFreeLabelTextLayout('1234\n123456789'), {
    characterCount: 13,
    columns: 8,
    rows: 3
  });
});

test('free label text cannot exceed 10 visual rows', () => {
  const constrained = constrainFreeLabelText(Array.from({ length: 12 }, () => '行').join('\n'));

  assert.equal(constrained.split('\n').length, FREE_LABEL_MAX_ROWS);
  assert.equal(getFreeLabelTextLayout(constrained).rows, FREE_LABEL_MAX_ROWS);
});

test('free label width contracts to the longest visible row', () => {
  assert.equal(getFreeLabelTextLayout('短い').columns, 2);
  assert.equal(getFreeLabelTextLayout('12345678').columns, 8);
});

test('free label formatting inserts a hard line break after every 8 characters', () => {
  assert.equal(formatFreeLabelText('123456789ABCDEFGH'), '12345678\n9ABCDEFG\nH');
  assert.equal(formatFreeLabelText('1234\n123456789'), '1234\n12345678\n9');
});
