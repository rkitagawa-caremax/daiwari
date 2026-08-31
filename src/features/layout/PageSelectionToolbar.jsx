import React from 'react';
import { ArrowLeftRight, CheckSquare, FileDown, Trash2, X } from 'lucide-react';

// 全体表示の「選択モード」中にヘッダーへ出す一括操作ツールバー。
const PageSelectionToolbar = ({
  selectedCount,
  totalCount,
  isProcessing,
  onSelectAll,
  onExportPdf,
  onSwap,
  onClearImages,
  onDelete
}) => (
  <div className="flex items-center gap-2 animate-in fade-in slide-in-from-left-4 duration-300 bg-white/50 backdrop-blur-sm px-2 py-1 rounded-xl border border-slate-200/50">
    <button
      onClick={onSelectAll}
      className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-bold whitespace-nowrap bg-white text-indigo-600 border border-indigo-100 hover:bg-indigo-50 hover:shadow-md`}
    >
      <CheckSquare size={14} />
      {selectedCount === totalCount && totalCount > 0 ? '全解除' : '全選択'}
    </button>
    <span className="text-sm font-bold text-slate-600 ml-1 mr-2 whitespace-nowrap">
      {selectedCount} / {totalCount}
    </span>
    <button
      onClick={onExportPdf}
      disabled={selectedCount === 0 || isProcessing}
      className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedCount > 0 && !isProcessing ? 'bg-emerald-600 text-white hover:bg-emerald-700 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
      title="選択したページを1つのPDFとして出力"
      data-work-action="pdf_export"
    >
      <FileDown size={14} /> PDF出力
    </button>
    <button
      onClick={onSwap}
      disabled={selectedCount !== 2}
      className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedCount === 2 ? 'bg-indigo-500 text-white hover:bg-indigo-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
      title="選択した2つのページを入れ替え"
    >
      <ArrowLeftRight size={14} /> 入れ替え
    </button>
    <button
      onClick={onClearImages}
      disabled={selectedCount === 0}
      className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedCount > 0 ? 'bg-amber-500 text-white hover:bg-amber-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
      title="選択したページの画像を全て外す"
    >
      <X size={14} /> 画像解除
    </button>
    <button
      onClick={onDelete}
      disabled={selectedCount === 0}
      className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedCount > 0 ? 'bg-rose-500 text-white hover:bg-rose-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
    >
      <Trash2 size={14} /> 削除
    </button>
  </div>
);

export default PageSelectionToolbar;
