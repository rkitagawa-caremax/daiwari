import { env, pipeline } from '@huggingface/transformers';

const MODEL_ID = 'ruri-v3-30m';

env.allowLocalModels = true;
env.allowRemoteModels = false;
env.localModelPath = '/models/';
env.useBrowserCache = true;
env.backends.onnx.wasm.wasmPaths = {
  mjs: '/models/wasm/ort-wasm-simd-threaded.mjs',
  wasm: '/models/wasm/ort-wasm-simd-threaded.wasm'
};
env.backends.onnx.wasm.numThreads = 1;

let extractorPromise = null;
let activeDevice = 'wasm';

const postProgress = (requestId, progress) => {
  self.postMessage({ type: 'progress', requestId, progress });
};

const createExtractor = async (requestId) => {
  const options = {
    dtype: 'int8',
    progress_callback: (progress) => postProgress(requestId, progress)
  };
  activeDevice = 'wasm';
  return pipeline('feature-extraction', MODEL_ID, { ...options, device: 'wasm' });
};

const getExtractor = (requestId) => {
  if (!extractorPromise) {
    extractorPromise = createExtractor(requestId).catch((error) => {
      extractorPromise = null;
      throw error;
    });
  }
  return extractorPromise;
};

self.addEventListener('message', async (event) => {
  const { requestId, type, texts = [] } = event.data || {};
  if (!requestId) return;
  try {
    const extractor = await getExtractor(requestId);
    if (type === 'init') {
      self.postMessage({ type: 'result', requestId, result: { device: activeDevice } });
      return;
    }
    if (type !== 'embed') throw new Error(`未対応のAI処理です: ${type}`);
    const safeTexts = (Array.isArray(texts) ? texts : []).map((text) => String(text || '').slice(0, 2400));
    if (!safeTexts.length) {
      self.postMessage({ type: 'result', requestId, result: { vectors: [], device: activeDevice } });
      return;
    }
    const output = await extractor(safeTexts, {
      pooling: 'mean',
      normalize: true,
      truncation: true,
      max_length: 512
    });
    self.postMessage({
      type: 'result',
      requestId,
      result: { vectors: output.tolist(), device: activeDevice }
    });
  } catch (error) {
    self.postMessage({
      type: 'error',
      requestId,
      error: error instanceof Error ? error.message : String(error)
    });
  }
});
