import React from 'react';
import {
  Link as LinkIcon,
  Merge,
  SlidersHorizontal,
  Split,
  Tag,
  Trash2
} from 'lucide-react';

const primaryButtonClass = 'flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left text-[11px] font-bold transition-colors disabled:cursor-not-allowed disabled:opacity-35';

const SheetControlPanel = React.memo(({
  viewMode,
  isLocked,
  isPageSelectionMode,
  isMergeMode,
  canMerge,
  canSplit,
  isLabelSelectionMode,
  activeSheetLabelCount,
  isPanelArrangeMode,
  onToggleMergeMode,
  onMerge,
  onSplit,
  onToggleLabelMode,
  onDeleteLabels,
  onShowQuickHelp,
  onHideQuickHelp
}) => {
  const isDetailView = viewMode === 'list' || viewMode === 'single';
  const isSinglePage = viewMode === 'single';
  const canUseMergeMode = isDetailView && !isPageSelectionMode && !isLocked;
  const canUseLabelTools = isSinglePage
    && !isPageSelectionMode
    && !isLocked
    && !isPanelArrangeMode;
  const canDeleteLabels = canUseLabelTools && activeSheetLabelCount > 0;

  return (
    <aside
      data-sheet-control-panel="true"
      className="fixed right-3 top-1/2 z-[92] w-40 -translate-y-1/2 rounded-2xl border border-slate-200/80 bg-white/90 p-2 shadow-lg shadow-slate-300/25 backdrop-blur-md"
      aria-label="ページ編集コントロール"
    >
      <div className="mb-1.5 flex items-center gap-2 px-1.5 py-1 text-[10px] font-bold tracking-wide text-slate-400">
        <SlidersHorizontal size={13} />
        コントロール
      </div>

      <button
        type="button"
        onClick={onToggleMergeMode}
        disabled={!canUseMergeMode}
        onMouseEnter={(event) => onShowQuickHelp?.(event, 'コマ結合', '複数コマを選択して結合・分離します。')}
        onMouseLeave={onHideQuickHelp}
        className={`${primaryButtonClass} ${isMergeMode && canUseMergeMode
          ? 'bg-indigo-50 text-indigo-700'
          : 'text-slate-600 hover:bg-slate-100'}`}
        title={canUseMergeMode ? 'コマ結合・分離モード' : '詳細表示で使用できます'}
        aria-pressed={isMergeMode}
      >
        <LinkIcon size={15} />
        <span>コマ結合</span>
      </button>

      {isMergeMode && canUseMergeMode && (
        <div className="mb-1 grid grid-cols-2 gap-1 px-1 pb-1 pt-1">
          <button
            type="button"
            onClick={onMerge}
            disabled={!canMerge}
            className="flex items-center justify-center gap-1 rounded-lg bg-emerald-50 px-1.5 py-1.5 text-[10px] font-bold text-emerald-700 transition-colors hover:bg-emerald-100 disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-slate-300"
            title="選択したコマを結合"
          >
            <Merge size={12} /> 結合
          </button>
          <button
            type="button"
            onClick={onSplit}
            disabled={!canSplit}
            className="flex items-center justify-center gap-1 rounded-lg bg-amber-50 px-1.5 py-1.5 text-[10px] font-bold text-amber-700 transition-colors hover:bg-amber-100 disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-slate-300"
            title="選択したコマを分離"
          >
            <Split size={12} /> 分離
          </button>
        </div>
      )}

      <div className="my-1 h-px bg-slate-200/70" />

      <button
        type="button"
        onClick={onToggleLabelMode}
        disabled={!canUseLabelTools}
        onMouseEnter={(event) => onShowQuickHelp?.(event, 'ラベル追加', 'コマ内をクリックして自由ラベルを配置します。')}
        onMouseLeave={onHideQuickHelp}
        className={`${primaryButtonClass} ${isLabelSelectionMode && canUseLabelTools
          ? 'bg-emerald-50 text-emerald-700'
          : 'text-slate-600 hover:bg-slate-100'}`}
        title={canUseLabelTools ? 'ラベル追加モード' : 'ページの詳細表示で使用できます'}
        aria-pressed={isLabelSelectionMode}
      >
        <Tag size={15} className={isLabelSelectionMode ? 'animate-pulse' : ''} />
        <span>{isLabelSelectionMode ? 'ラベル配置中' : 'ラベル追加'}</span>
      </button>

      <button
        type="button"
        onClick={onDeleteLabels}
        disabled={!canDeleteLabels}
        onMouseEnter={(event) => onShowQuickHelp?.(event, 'ラベル一括削除', '表示中のページにある自由ラベルをまとめて削除します。')}
        onMouseLeave={onHideQuickHelp}
        className={`${primaryButtonClass} text-slate-600 hover:bg-rose-50 hover:text-rose-700`}
        title={canDeleteLabels ? 'このページのラベルを一括削除' : '削除できるラベルがありません'}
      >
        <Trash2 size={15} />
        <span className="min-w-0 flex-1">ラベル一括削除</span>
        {activeSheetLabelCount > 0 && isSinglePage && (
          <span className="rounded-full bg-slate-100 px-1.5 py-0.5 font-mono text-[9px] text-slate-500">
            {activeSheetLabelCount}
          </span>
        )}
      </button>
    </aside>
  );
});

SheetControlPanel.displayName = 'SheetControlPanel';

export default SheetControlPanel;
