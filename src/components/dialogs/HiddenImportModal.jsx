import React from 'react';
import { BarChart3, FileText, TrendingUp, Upload, X } from 'lucide-react';

const HiddenImportModal = React.memo(({
  isOpen,
  onClose,
  onOpenPageCsvImport,
  onOpenSalesCsvImport,
  onOpenWorkLogs
}) => {
  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-[85] flex items-center justify-center bg-black/45 backdrop-blur-sm m3-animate-fade-in"
      onClick={onClose}
    >
      <div
        className="m3-dialog w-[420px] overflow-hidden p-0 m3-animate-scale-in"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="p-5 border-b flex justify-between items-center" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <h3 className="text-lg font-medium flex items-center gap-2" style={{ color: 'var(--m3-on-surface)' }}>
            <Upload size={18} />
            管理メニュー
          </h3>
          <button onClick={onClose} className="m3-icon-btn">
            <X size={18} />
          </button>
        </div>

        <div className="p-5 space-y-3" style={{ background: 'var(--m3-surface-container-high)' }}>
          <button
            onClick={onOpenPageCsvImport}
            className="w-full m3-btn-tonal flex items-center justify-center gap-2"
          >
            <FileText size={16} />
            台割CSVを取り込む
          </button>

          <button
            onClick={onOpenSalesCsvImport}
            className="w-full m3-btn-tonal flex items-center justify-center gap-2"
          >
            <TrendingUp size={16} />
            販売数量CSVを取り込む
          </button>

          <button
            onClick={onOpenWorkLogs}
            className="w-full m3-btn-tonal flex items-center justify-center gap-2"
            data-work-action="settings"
          >
            <BarChart3 size={16} />
            ログ
          </button>
        </div>
      </div>
    </div>
  );
});

export default HiddenImportModal;
