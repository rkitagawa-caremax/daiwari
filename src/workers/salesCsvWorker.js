import { parseSalesCsvContent } from '../domain/salesData.js';

self.onmessage = (event) => {
  try {
    const salesData = parseSalesCsvContent(event.data);
    self.postMessage({ ok: true, salesData });
  } catch (error) {
    self.postMessage({
      ok: false,
      message: error instanceof Error ? error.message : String(error)
    });
  }
};

