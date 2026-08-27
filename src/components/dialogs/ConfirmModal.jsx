import React from 'react';
import { AlertTriangle } from 'lucide-react';

const ConfirmModal = React.memo(({ isOpen, message, onConfirm, onCancel }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in">
      <div className="m3-dialog m3-animate-scale-in" style={{ background: 'var(--m3-surface-container-high)' }}>
        <div className="flex items-center gap-4 mb-6">
          <div className="p-3 rounded-full" style={{ background: 'var(--m3-error-container)' }}>
            <AlertTriangle size={24} style={{ color: 'var(--m3-on-error-container)' }} />
          </div>
          <h3 className="text-xl font-medium" style={{ color: 'var(--m3-on-surface)' }}>確認</h3>
        </div>
        <p className="text-sm mb-8 whitespace-pre-wrap leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>
        <div className="flex justify-end gap-3">
          <button onClick={onCancel} className="m3-btn-text">
            キャンセル
          </button>
          <button
            onClick={onConfirm}
            className="px-6 py-2.5 font-medium transition-all active:scale-[0.98]"
            style={{
              background: 'var(--m3-error)',
              color: 'var(--m3-on-error)',
              borderRadius: 'var(--m3-shape-corner-full)'
            }}
          >
            実行する
          </button>
        </div>
      </div>
    </div>
  );
});

export default ConfirmModal;
