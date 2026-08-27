import React, { useRef } from 'react';
import { Database, FileText, Settings, TrendingUp, X } from 'lucide-react';

const SettingsModal = React.memo(({ isOpen, onClose, onImportSalesCSV, salesDataLastUpdated }) => {
  const fileInputRef = useRef(null);

  if (!isOpen) return null;

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) {
      onImportSalesCSV(file);
    }
  };

  return (
    <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in">
      <div className="m3-dialog w-[520px] overflow-hidden flex flex-col max-h-[90vh] m3-animate-scale-in p-0" style={{ padding: 0 }}>
        <div className="p-5 border-b flex justify-between items-center" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <h3 className="text-lg font-medium flex items-center gap-3" style={{ color: 'var(--m3-on-surface)' }}>
            <div className="p-2 rounded-full" style={{ background: 'var(--m3-secondary-container)' }}>
              <Settings className="w-5 h-5" style={{ color: 'var(--m3-on-secondary-container)' }} />
            </div>
            設定
          </h3>
          <button onClick={onClose} className="m3-icon-btn">
            <X size={20} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-8" style={{ background: 'var(--m3-surface-container-high)' }}>
          <section>
            <h4 className="text-sm font-medium mb-4 flex items-center gap-3" style={{ color: 'var(--m3-on-surface)' }}>
              <div className="p-1.5 rounded-full" style={{ background: 'var(--m3-tertiary-container)' }}>
                <TrendingUp className="w-4 h-4" style={{ color: 'var(--m3-on-tertiary-container)' }} />
              </div>
              販売数量データの取り込み
            </h4>
            <div className="p-5" style={{ background: 'var(--m3-surface-container-lowest)', borderRadius: 'var(--m3-shape-corner-lg)' }}>
              <p className="text-sm mb-5 leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>
                CSVファイル（商品別売上推移表）を取り込むと、パネル上のコード（介援隊CD）と照合して販売数量を表示できます。<br />
                <span style={{ color: 'var(--m3-error)' }}>※ 取り込みを行うと以前のデータは上書きされます。</span>
              </p>

              <div className="flex items-center gap-4 mb-4">
                <input
                  type="file"
                  accept=".csv"
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="m3-btn-tonal flex items-center gap-2"
                >
                  <FileText size={18} /> CSVを選択して取り込む
                </button>
              </div>

              {salesDataLastUpdated && (
                <div className="flex items-center gap-2 text-xs font-mono px-3 py-2 w-fit" style={{ background: 'var(--m3-surface-container)', borderRadius: 'var(--m3-shape-corner-sm)', color: 'var(--m3-outline)' }}>
                  <Database size={12} />
                  最終更新: {new Date(salesDataLastUpdated).toLocaleString()}
                </div>
              )}
            </div>
          </section>
        </div>

        <div className="p-4 border-t flex justify-end" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <button onClick={onClose} className="m3-btn-outlined">
            閉じる
          </button>
        </div>
      </div>
    </div>
  );
});

export default SettingsModal;
