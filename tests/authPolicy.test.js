import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildGoogleAuthErrorMessage,
  expandAllowedEmailVariants,
  isAllowedGoogleUser,
  normalizeEmail
} from '../src/config/authPolicy.js';

test('allowed Google accounts accept normalized domain variants', () => {
  assert.equal(normalizeEmail('  Y.GOTO@G.CAREMAX.CO.JP  '), 'y.goto@g.caremax.co.jp');
  assert.deepEqual(
    expandAllowedEmailVariants('y.goto@g.caremax.co.jp'),
    ['y.goto@g.caremax.co.jp', 'y.goto@caremax.co.jp']
  );
  assert.equal(isAllowedGoogleUser({ email: 'Y.GOTO@caremax.co.jp' }), true);
  assert.equal(isAllowedGoogleUser({ email: 'unknown@example.com' }), false);
  assert.equal(isAllowedGoogleUser(null), false);
});

test('Google authentication errors retain their user-facing messages', () => {
  assert.equal(
    buildGoogleAuthErrorMessage({ code: 'auth/popup-closed-by-user' }),
    'ログインがキャンセルされました。'
  );
  assert.equal(
    buildGoogleAuthErrorMessage({ code: 'auth/network-request-failed' }),
    'ネットワークエラーでログインできませんでした。接続を確認して再試行してください。'
  );
  assert.equal(buildGoogleAuthErrorMessage({ code: 'unknown' }, 'fallback'), 'fallback');
});
