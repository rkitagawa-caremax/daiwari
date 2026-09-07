import { parseSalesCsvContent } from '../domain/salesData.js';

self.onmessage = (event) => {
  try {
    const payload = typeof event.data === 'string' ? { csvText: event.data, options: {} } : event.data;
    const salesData = parseSalesCsvContent(payload?.csvText, payload?.options);
    self.postMessage({ ok: true, salesData });
  } catch (error) {
    self.postMessage({
      ok: false,
      message: error instanceof Error ? error.message : String(error)
    });
  }
};
