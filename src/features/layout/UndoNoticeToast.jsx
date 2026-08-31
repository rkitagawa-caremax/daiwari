import React from 'react';

// 戻る/進む 実行結果のトースト (画面下中央)。notice = { message, tone: 'success' | 'warning' | 'neutral' }
const UndoNoticeToast = ({ notice }) => {
  if (!notice) return null;
  return (
    <div
      className={`fixed bottom-5 left-1/2 z-[210] -translate-x-1/2 rounded-full border px-4 py-2 text-xs font-bold shadow-lg backdrop-blur-md ${notice.tone === 'warning'
        ? 'border-amber-200 bg-amber-50/95 text-amber-800'
        : notice.tone === 'neutral'
          ? 'border-slate-200 bg-white/95 text-slate-600'
          : 'border-emerald-200 bg-emerald-50/95 text-emerald-800'
        }`}
      role="status"
      aria-live="polite"
    >
      {notice.message}
    </div>
  );
};

export default UndoNoticeToast;
