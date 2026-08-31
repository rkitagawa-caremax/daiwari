import React from 'react';
import { ZoomIn, ZoomOut } from 'lucide-react';

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 1.5;
export const ZOOM_STEP = 0.05;
export const DEFAULT_ZOOM_SCALE = 0.85;

// 右下に固定表示する表示倍率コントロール (詳細表示のみ)。
const ZoomControls = ({ zoomScale, setZoomScale }) => (
  <div
    className="fixed bottom-4 right-4 z-[90] flex items-center gap-0.5 rounded-xl border border-slate-200/80 bg-white/80 p-1 text-slate-500 shadow-sm backdrop-blur opacity-65 transition-all duration-200 hover:bg-white/95 hover:opacity-100 hover:shadow-md focus-within:opacity-100"
    aria-label="表示倍率"
  >
    <button
      type="button"
      onClick={() => setZoomScale((scale) => Math.max(ZOOM_MIN, scale - ZOOM_STEP))}
      disabled={zoomScale <= ZOOM_MIN}
      className="rounded-lg p-1.5 transition-colors hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-30"
      title="縮小"
      aria-label="表示を縮小"
    >
      <ZoomOut size={15} />
    </button>
    <span className="w-10 select-none text-center font-mono text-[10px] font-bold text-slate-500" aria-live="polite">
      {Math.round(zoomScale * 100)}%
    </span>
    <button
      type="button"
      onClick={() => setZoomScale((scale) => Math.min(ZOOM_MAX, scale + ZOOM_STEP))}
      disabled={zoomScale >= ZOOM_MAX}
      className="rounded-lg p-1.5 transition-colors hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-30"
      title="拡大"
      aria-label="表示を拡大"
    >
      <ZoomIn size={15} />
    </button>
  </div>
);

export default ZoomControls;
