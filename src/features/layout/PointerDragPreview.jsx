import React from 'react';

// ポインタ DnD 中にカーソルへ追従するプレビュー。
// 位置は App 側が ref 経由で transform を直接書き換える (再レンダリングを避けるため) ので ref を受け取る。
const PointerDragPreview = React.forwardRef(({ preview }, ref) => {
  if (!preview) return null;
  return (
    <div
      ref={ref}
      className="fixed left-0 top-0 z-[160] pointer-events-none will-change-transform"
    >
      <div className="min-w-[96px] max-w-[144px] rounded-2xl border border-sky-200 bg-white/95 p-2 shadow-2xl backdrop-blur-md">
        {preview.image ? (
          <div className="aspect-square w-24 overflow-hidden rounded-xl bg-slate-100 flex items-center justify-center">
            <img
              src={preview.image}
              alt="drag preview"
              className="max-w-full max-h-full object-contain"
              draggable={false}
            />
          </div>
        ) : (
          <div className="flex h-20 w-24 items-center justify-center rounded-xl bg-slate-100 px-2 text-center text-xs font-bold text-slate-600">
            {preview.label || preview.code || (preview.text ? 'テキスト' : '移動')}
          </div>
        )}
        <p className="mt-1.5 truncate text-center text-[10px] font-bold text-slate-700">
          {preview.code || preview.label || preview.text || '移動中'}
        </p>
      </div>
    </div>
  );
});

PointerDragPreview.displayName = 'PointerDragPreview';

export default PointerDragPreview;
