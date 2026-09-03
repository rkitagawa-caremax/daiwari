let worker = null;
let requestSequence = 0;
const pendingRequests = new Map();

const rejectPendingRequests = (message) => {
  pendingRequests.forEach(({ reject }) => reject(new Error(message)));
  pendingRequests.clear();
};

const getWorker = () => {
  if (worker) return worker;
  worker = new Worker(new URL('../workers/edgeAiWorker.js', import.meta.url), {
    type: 'module',
    name: 'daiwari-edge-ai'
  });
  worker.addEventListener('message', (event) => {
    const { requestId, type, result, error, progress } = event.data || {};
    const request = pendingRequests.get(requestId);
    if (!request) return;
    if (type === 'progress') {
      request.onProgress?.(progress);
      return;
    }
    pendingRequests.delete(requestId);
    if (type === 'error') request.reject(new Error(error || '端末内AIの処理に失敗しました。'));
    else request.resolve(result);
  });
  worker.addEventListener('error', (event) => {
    rejectPendingRequests(event.message || '端末内AIを起動できませんでした。');
    worker?.terminate();
    worker = null;
  });
  return worker;
};

const callWorker = (type, payload = {}, { onProgress } = {}) => new Promise((resolve, reject) => {
  const requestId = `edge-ai-${Date.now()}-${requestSequence += 1}`;
  pendingRequests.set(requestId, { resolve, reject, onProgress });
  getWorker().postMessage({ requestId, type, ...payload });
});

export const initializeEdgeAi = (options) => callWorker('init', {}, options);

export const embedEdgeAiTexts = (texts, options) => callWorker('embed', { texts }, options);

export const disposeEdgeAiWorker = () => {
  if (!worker) return;
  rejectPendingRequests('端末内AIを終了しました。');
  worker.terminate();
  worker = null;
};

