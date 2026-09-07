import { parseSalesCsvContent } from '../domain/salesData.js';

// 大きな売上CSVでも画面操作を止めないよう、対応ブラウザでは解析を専用スレッドへ逃がす。
export const parseSalesCsvWithoutBlocking = (csvText, options = {}) => {
  if (typeof Worker === 'undefined') {
    return Promise.resolve(parseSalesCsvContent(csvText, options));
  }

  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('../workers/salesCsvWorker.js', import.meta.url), {
      type: 'module',
      name: 'daiwari-sales-csv-parser'
    });

    const finish = () => worker.terminate();
    worker.addEventListener('message', (event) => {
      finish();
      if (event.data?.ok) {
        resolve(event.data.salesData || {});
      } else {
        reject(new Error(event.data?.message || '売上CSVの解析に失敗しました。'));
      }
    }, { once: true });
    worker.addEventListener('error', (event) => {
      finish();
      reject(new Error(event.message || '売上CSVの解析処理を開始できませんでした。'));
    }, { once: true });
    worker.postMessage({ csvText, options });
  });
};
