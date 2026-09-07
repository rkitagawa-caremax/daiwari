import React, { useRef, useState } from 'react';
import { BadgeJapaneseYen, CircleDollarSign, Database, PackageCheck, Settings, TrendingUp, X } from 'lucide-react';

const formatUpdatedAt = (value) => {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value?.seconds ? value.seconds * 1000 : value);
  return Number.isNaN(date.getTime()) ? '' : date.toLocaleString();
};

// 期 (今期 / 前期 / 前々期) ごとに販売実績CSVを取り込む。
// 期を選んでから CSV を選ぶ流れにして、どの期に入るかを取り込み前に確かめられるようにしている。
const SettingsModal = React.memo(({
  isOpen,
  onClose,
  onImportSalesCSV,
  salesPeriodOptions = [],
  salesPeriodMeta = {}
}) => {
  const fileInputRefs = useRef({});
  const [targetPeriodId, setTargetPeriodId] = useState(salesPeriodOptions[0]?.id || 'current');

  if (!isOpen) return null;

  const targetPeriod = salesPeriodOptions.find((period) => period.id === targetPeriodId) || salesPeriodOptions[0];

  const handleFileChange = (event, metricType) => {
    const file = event.target.files[0];
    event.target.value = '';
    if (file) onImportSalesCSV(file, targetPeriodId, metricType);
  };

  const importOptions = [
    { id: 'quantity', label: '販売数量CSV', note: '売れた個数', icon: PackageCheck, cardClass: 'border-violet-100 bg-violet-50/60', iconClass: 'text-violet-600' },
    { id: 'salesAmount', label: '売上額CSV', note: '販売金額', icon: CircleDollarSign, cardClass: 'border-sky-100 bg-sky-50/60', iconClass: 'text-sky-600' },
    { id: 'grossProfitAmount', label: '粗利額CSV', note: '粗利益額', icon: BadgeJapaneseYen, cardClass: 'border-emerald-100 bg-emerald-50/60', iconClass: 'text-emerald-600' }
  ];

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
              販売実績データの取り込み
            </h4>
            <div className="p-5" style={{ background: 'var(--m3-surface-container-lowest)', borderRadius: 'var(--m3-shape-corner-lg)' }}>
              <p className="text-sm mb-4 leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>
                販売数量・売上額・粗利額を別々のCSVから取り込み、介援隊コードで1つの実績へ統合します。
                取り込んだ指標だけを更新するため、ほかの実績は消えません。
              </p>

              <div className="mb-4 rounded-2xl border border-emerald-100 bg-emerald-50/70 px-4 py-3">
                <p className="flex items-center gap-2 text-xs font-bold text-emerald-800"><CircleDollarSign size={15} /> 自動認識する主な列</p>
                <p className="mt-1.5 text-[11px] leading-5 text-emerald-700">介援隊コード／商品コードと、選択した実績列を照合します。列名が独自形式でも、各CSVの18列目を選択した指標として取り込めます。</p>
              </div>

              <p className="text-xs font-bold mb-2" style={{ color: 'var(--m3-on-surface-variant)' }}>取り込む期を選ぶ</p>
              <div className="space-y-2 mb-4">
                {salesPeriodOptions.map((period) => {
                  const meta = salesPeriodMeta?.[period.id];
                  const metricLabels = [
                    meta?.metrics?.quantity?.codes > 0 ? '数量' : '',
                    meta?.metrics?.salesAmount?.codes > 0 ? '売上額' : '',
                    meta?.metrics?.grossProfitAmount?.codes > 0 ? '粗利額' : ''
                  ].filter(Boolean);
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
                            ? `${(meta.totalItems || 0).toLocaleString()}商品${metricLabels.length ? ` / ${metricLabels.join('・')}` : ' / 旧形式の数量データ'}${updatedAt ? ` / 最終更新 ${updatedAt}` : ''}${meta.fileName ? ` / ${meta.fileName}` : ''}`
                            : 'まだ取り込まれていません'}
                        </span>
                      </span>
                    </label>
                  );
                })}
              </div>

              <p className="mb-2 text-xs font-bold" style={{ color: 'var(--m3-on-surface-variant)' }}>{targetPeriod?.label || ''}へ取り込むデータを選ぶ</p>
              <div className="grid grid-cols-3 gap-2">
                {importOptions.map((option) => {
                  const Icon = option.icon;
                  const sourceFile = salesPeriodMeta?.[targetPeriodId]?.sourceFiles?.[option.id];
                  return (
                    <div key={option.id} className={`rounded-2xl border p-2.5 ${option.cardClass}`}>
                      <input
                        type="file"
                        accept=".csv"
                        ref={(node) => { fileInputRefs.current[option.id] = node; }}
                        onChange={(event) => handleFileChange(event, option.id)}
                        className="hidden"
                      />
                      <button
                        onClick={() => fileInputRefs.current[option.id]?.click()}
                        className="flex w-full flex-col items-center rounded-xl bg-white px-2 py-3 text-center shadow-sm transition hover:-translate-y-0.5 hover:shadow-md"
                      >
                        <Icon size={21} className={`mb-1.5 ${option.iconClass}`} />
                        <span className="text-xs font-bold text-slate-800">{option.label}</span>
                        <span className="mt-0.5 text-[10px] text-slate-500">{option.note}</span>
                      </button>
                      {sourceFile && <p className="mt-1.5 truncate text-center text-[9px] text-slate-500" title={sourceFile}>{sourceFile}</p>}
                    </div>
                  );
                })}
              </div>

              {salesPeriodMeta?.[targetPeriodId] && (
                <div className="mt-4 flex items-center gap-2 text-xs px-3 py-2 w-fit" style={{ background: 'var(--m3-surface-container)', borderRadius: 'var(--m3-shape-corner-sm)', color: 'var(--m3-error)' }}>
                  <Database size={12} />
                  選んだ指標のみ更新し、ほかの実績は保持します
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
