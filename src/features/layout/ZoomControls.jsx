import React from 'react';
import { Minus, Plus } from 'lucide-react';

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 1.5;
export const ZOOM_STEP = 0.05;
export const DEFAULT_ZOOM_SCALE = 0.85;

// 右下に固定表示する表示倍率コントロール (詳細表示のみ)。
const ZoomControls = ({ zoomScale, setZoomScale }) => (
  <div
    className="fixed bottom-4 right-4 z-[90] flex h-10 items-center gap-2 rounded-full border border-slate-200 bg-white px-2.5 text-slate-500 shadow-[0_3px_12px_rgba(15,23,42,0.12)]"
    aria-label="表示倍率"
  >
    <button
      type="button"
      onClick={() => setZoomScale((scale) => Math.max(ZOOM_MIN, scale - ZOOM_STEP))}
      disabled={zoomScale <= ZOOM_MIN}
      className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full transition-colors hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-30"
      title="縮小"
      aria-label="表示を縮小"
    >
      <Minus size={17} strokeWidth={2.25} />
    </button>
    <input
      type="range"
      min={ZOOM_MIN}
      max={ZOOM_MAX}
      step={ZOOM_STEP}
      value={zoomScale}
      onChange={(event) => setZoomScale(Number(event.target.value))}
      className="daiwari-detail-zoom-range w-28 sm:w-32"
      style={{ '--daiwari-zoom-progress': `${((zoomScale - ZOOM_MIN) / (ZOOM_MAX - ZOOM_MIN)) * 100}%` }}
      aria-label="表示倍率を変更"
      aria-valuetext={`${Math.round(zoomScale * 100)}%`}
    />
    <button
      type="button"
      onClick={() => setZoomScale((scale) => Math.min(ZOOM_MAX, scale + ZOOM_STEP))}
      disabled={zoomScale >= ZOOM_MAX}
      className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full transition-colors hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-30"
      title="拡大"
      aria-label="表示を拡大"
    >
      <Plus size={17} strokeWidth={2.25} />
    </button>
    <span className="min-w-14 select-none rounded-full bg-slate-100 px-2.5 py-1 text-center font-mono text-[10px] font-bold text-slate-500" aria-live="polite">
      {Math.round(zoomScale * 100)}%
    </span>
  </div>
);

export default ZoomControls;
