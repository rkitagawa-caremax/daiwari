import { useEffect } from 'react';

const ImagePreviewModal = ({ preview, onClose }) => {
  useEffect(() => {
    if (!preview?.src) return undefined;

    const handleKeyDown = (event) => {
      if (event.key === 'Escape') onClose?.();
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [preview, onClose]);

  if (!preview?.src) return null;

  return (
    <div
      className="fixed inset-0 z-[170] flex items-center justify-center bg-black/70 p-5"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-label="画像プレビュー"
      title="クリックまたはEscキーでプレビューを閉じる"
    >
      <div className="max-h-[90vh] max-w-[92vw] rounded-2xl border border-white/20 bg-slate-950/90 p-3 shadow-2xl">
        <img
          src={preview.src}
          alt={preview.name || 'preview'}
          className="max-h-[80vh] max-w-[86vw] rounded-lg object-contain"
        />
        {preview.name && (
          <p className="mt-2 truncate text-center font-mono text-[11px] text-slate-200">
            {preview.name}
          </p>
        )}
      </div>
    </div>
  );
};

export default ImagePreviewModal;
