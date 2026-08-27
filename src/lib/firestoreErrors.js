const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export const getFirestoreErrorCode = (error) => {
  if (!error) return '';
  const raw = String(error.code || '').toLowerCase();
  return raw.startsWith('firestore/') ? raw.replace('firestore/', '') : raw;
};

export const buildFirestoreActionErrorMessage = (fallbackMessage, error) => {
  const code = getFirestoreErrorCode(error);
  if (code === 'resource-exhausted') {
    return `${fallbackMessage}\n\nFirestoreの無料枠またはクォータに達している可能性があります。Firebaseコンソールの使用状況を確認してください。`;
  }
  if (code === 'permission-denied') {
    return `${fallbackMessage}\n\nFirestoreの権限が不足している可能性があります。ログインアカウントとFirestoreルールを確認してください。`;
  }
  if (code === 'failed-precondition' || code === 'aborted') {
    return `${fallbackMessage}\n\nFirestoreの同時編集競合が発生しました。少し待ってから再実行してください。`;
  }
  if (code === 'invalid-argument') {
    return `${fallbackMessage}\n\nFirestoreに保存できないデータ形式またはサイズの可能性があります。`;
  }
  return fallbackMessage;
};

const RETRYABLE_FIRESTORE_CODES = new Set([
  'aborted',
  'cancelled',
  'deadline-exceeded',
  'failed-precondition',
  'internal',
  'resource-exhausted',
  'unavailable',
  'unknown'
]);

export const isRetryableFirestoreError = (error) => (
  RETRYABLE_FIRESTORE_CODES.has(getFirestoreErrorCode(error))
);

export const retryAsync = async (task, options = {}) => {
  const {
    retries = 8,
    baseDelayMs = 80,
    maxDelayMs = 1600
  } = options;

  let attempt = 0;
  while (true) {
    try {
      return await task(attempt);
    } catch (error) {
      if (attempt >= retries || !isRetryableFirestoreError(error)) {
        throw error;
      }
      const backoff = Math.min(baseDelayMs * (2 ** attempt), maxDelayMs);
      const jitter = Math.floor(Math.random() * 60);
      await sleep(backoff + jitter);
      attempt += 1;
    }
  }
};
