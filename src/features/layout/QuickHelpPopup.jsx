import React from 'react';

// クイックヘルプ (Q モード) のツールチップ。表示可否は呼び出し側で判定する。
const QuickHelpPopup = ({ popup }) => {
  if (!popup) return null;
  return (
    <div
      className="fixed z-[120] pointer-events-none"
      style={{ left: popup.x, top: popup.y, transform: 'translateX(-50%)' }}
    >
      <div className="min-w-[320px] max-w-[460px] rounded-2xl border border-sky-200 bg-white/95 backdrop-blur-sm px-4 py-3 shadow-xl">
        <p className="text-[13px] font-bold text-sky-700">{popup.title}</p>
        <p className="text-[12px] leading-relaxed text-slate-700 mt-1.5">{popup.description}</p>
      </div>
    </div>
  );
};

export default QuickHelpPopup;
