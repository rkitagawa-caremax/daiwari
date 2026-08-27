import React from 'react';
import { Bell } from 'lucide-react';

const AlertModal = React.memo(({
  isOpen,
  message,
  onClose,
  title = '通知',
  closeOnBackdrop = false
}) => {
  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-[90] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in"
      onClick={() => {
        if (closeOnBackdrop) {
          onClose();
        }
      }}
    >
      <div
        className="m3-dialog m3-animate-scale-in"
        style={{ background: 'var(--m3-surface-container-high)' }}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-center gap-4 mb-6">
          <div className="p-3 rounded-full" style={{ background: 'var(--m3-primary-container)' }}>
            <Bell size={24} style={{ color: 'var(--m3-on-primary-container)' }} />
          </div>
          <h3 className="text-xl font-medium" style={{ color: 'var(--m3-on-surface)' }}>{title}</h3>
        </div>
        <p className="text-sm mb-8 whitespace-pre-wrap leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>
        <div className="flex justify-end">
          <button onClick={onClose} className="m3-btn-filled">
            OK
          </button>
        </div>
      </div>
    </div>
  );
});

export default AlertModal;
