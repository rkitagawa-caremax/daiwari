import React from 'react';
import { AlertTriangle, Loader2 } from 'lucide-react';

const ProcessingModal = React.memo(({ isOpen, current, total, message }) => {
  if (!isOpen) return null;

  const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/50 backdrop-blur-sm cursor-wait m3-animate-fade-in">
      <div className="m3-dialog w-96 flex flex-col items-center m3-animate-scale-in">
        <div className="relative mb-6">
          <div className="absolute inset-0 rounded-full blur-xl opacity-30 animate-pulse" style={{ background: 'var(--m3-primary)' }}></div>
          <Loader2 className="w-14 h-14 animate-spin relative z-10" style={{ color: 'var(--m3-primary)' }} />
        </div>
        <h3 className="text-xl font-medium mb-2" style={{ color: 'var(--m3-on-surface)' }}>処理中...</h3>
        <p className="text-sm mb-6 text-center whitespace-pre-wrap" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>

        <div className="w-full h-1 rounded-full mb-3 overflow-hidden" style={{ background: 'var(--m3-surface-container-highest)' }}>
          <div
            className="h-1 rounded-full transition-all duration-300 ease-out"
            style={{ width: `${percentage}%`, background: 'var(--m3-primary)' }}
          ></div>
        </div>
        <div className="flex justify-between w-full text-xs font-mono" style={{ color: 'var(--m3-outline)' }}>
          <span>{current} / {total}</span>
          <span>{percentage}%</span>
        </div>
        <p className="text-xs mt-6 font-medium flex items-center gap-2" style={{ color: 'var(--m3-error)' }}>
          <AlertTriangle size={14} /> 画面を閉じないでください
        </p>
      </div>
    </div>
  );
});

export default ProcessingModal;
