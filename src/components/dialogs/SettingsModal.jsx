import React, { useRef, useState } from 'react';
import { Database, FileText, Settings, TrendingUp, X } from 'lucide-react';

const formatUpdatedAt = (value) => {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value?.seconds ? value.seconds * 1000 : value);
  return Number.isNaN(date.getTime()) ? '' : date.toLocaleString();
};

// 期 (今期 / 前期 / 前々期) ごとに売上CSVを取り込む。
// 期を選んでから CSV を選ぶ流れにして、どの期に入るかを取り込み前に確かめられるようにしている。
const SettingsModal = React.memo(({
  isOpen,
  onClose,
  onImportSalesCSV,
  salesPeriodOptions = [],
  salesPeriodMeta = {}
}) => {
  const fileInputRef = useRef(null);
  const [targetPeriodId, setTargetPeriodId] = useState(salesPeriodOptions[0]?.id || 'current');

  if (!isOpen) return null;

  const targetPeriod = salesPeriodOptions.find((period) => period.id === targetPeriodId) || salesPeriodOptions[0];

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    event.target.value = '';
    if (file) onImportSalesCSV(file, targetPeriodId);
  };

  return (
    <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in">
      <div className="m3-dialog w-[560px] overflow-hidden flex flex-col max-h-[90vh] m3-animate-scale-in p-0" style={{ padding: 0 }}>
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
              <p className="text-sm mb-4 leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>
                CSVファイル（商品別売上推移表）を取り込むと、パネル上のコード（介援隊CD）と照合して販売数量と月別推移を表示できます。
                期ごとに保存でき、実績モードのヘッダーで切り替えて比べられます。
              </p>

              <p className="text-xs font-bold mb-2" style={{ color: 'var(--m3-on-surface-variant)' }}>取り込む期を選ぶ</p>
              <div className="space-y-2 mb-4">
                {salesPeriodOptions.map((period) => {
                  const meta = salesPeriodMeta?.[period.id];
                  const updatedAt = formatUpdatedAt(meta?.updatedAt);
                  const isSelected = period.id === targetPeriodId;
                  return (
                    <label
                      key={period.id}
                      className="flex cursor-pointer items-center gap-3 p-3 transition-colors"
                      style={{
                        borderRadius: 'var(--m3-shape-corner-md)',
                        background: isSelected ? 'var(--m3-secondary-container)' : 'var(--m3-surface-container)',
                        outline: isSelected ? '2px solid var(--m3-primary)' : 'none'
                      }}
                    >
                      <input
                        type="radio"
                        name="sales-period"
                        value={period.id}
                        checked={isSelected}
                        onChange={() => setTargetPeriodId(period.id)}
                        className="shrink-0"
                      />
                      <span className="min-w-0 flex-1">
                        <span className="block text-sm font-bold" style={{ color: 'var(--m3-on-surface)' }}>
                          {period.label}
                          <span className="ml-2 text-[11px] font-medium" style={{ color: 'var(--m3-outline)' }}>{period.description}</span>
                        </span>
                        <span className="block text-[11px] mt-0.5" style={{ color: 'var(--m3-on-surface-variant)' }}>
                          {meta
                            ? `${(meta.totalItems || 0).toLocaleString()}商品${updatedAt ? ` / 最終更新 ${updatedAt}` : ''}${meta.fileName ? ` / ${meta.fileName}` : ''}`
                            : 'まだ取り込まれていません'}
                        </span>
                      </span>
                    </label>
                  );
                })}
              </div>

              <div className="flex items-center gap-4">
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
                  <FileText size={18} /> {targetPeriod?.label || ''}にCSVを取り込む
                </button>
              </div>

              {salesPeriodMeta?.[targetPeriodId] && (
                <div className="mt-4 flex items-center gap-2 text-xs px-3 py-2 w-fit" style={{ background: 'var(--m3-surface-container)', borderRadius: 'var(--m3-shape-corner-sm)', color: 'var(--m3-error)' }}>
                  <Database size={12} />
                  取り込むと{targetPeriod?.label}の既存データは上書きされます
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
