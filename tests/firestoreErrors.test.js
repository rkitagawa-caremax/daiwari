import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildFirestoreActionErrorMessage,
  getFirestoreErrorCode,
  isRetryableFirestoreError,
  retryAsync
} from '../src/lib/firestoreErrors.js';

test('Firestore error codes are normalized before classification', () => {
  assert.equal(getFirestoreErrorCode({ code: 'firestore/Permission-Denied' }), 'permission-denied');
  assert.equal(getFirestoreErrorCode({ code: 'UNAVAILABLE' }), 'unavailable');
  assert.equal(getFirestoreErrorCode(null), '');
  assert.equal(isRetryableFirestoreError({ code: 'firestore/aborted' }), true);
  assert.equal(isRetryableFirestoreError({ code: 'permission-denied' }), false);
});

test('Firestore action messages retain the fallback and add known guidance', () => {
  const fallback = '更新に失敗しました。';
  assert.match(
    buildFirestoreActionErrorMessage(fallback, { code: 'resource-exhausted' }),
    /^更新に失敗しました。\n\nFirestoreの無料枠/
  );
  assert.equal(
    buildFirestoreActionErrorMessage(fallback, { code: 'not-found' }),
    fallback
  );
});

test('retryAsync retries retryable failures and stops at the configured limit', async () => {
  const originalRandom = Math.random;
  Math.random = () => 0;

  try {
    let attempts = 0;
    const result = await retryAsync(async (attempt) => {
      attempts += 1;
      if (attempt < 2) throw { code: 'firestore/unavailable' };
      return 'ok';
    }, { retries: 2, baseDelayMs: 0, maxDelayMs: 0 });

    assert.equal(result, 'ok');
    assert.equal(attempts, 3);

    let nonRetryableAttempts = 0;
    await assert.rejects(
      retryAsync(async () => {
        nonRetryableAttempts += 1;
        throw { code: 'permission-denied' };
      }, { retries: 8, baseDelayMs: 0, maxDelayMs: 0 }),
      (error) => error.code === 'permission-denied'
    );
    assert.equal(nonRetryableAttempts, 1);
  } finally {
    Math.random = originalRandom;
  }
});
