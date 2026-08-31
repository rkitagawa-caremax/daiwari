import React from 'react';
import { ArrowLeftRight, Check, Loader2 } from 'lucide-react';

// 画像ホバリング (パネル配置モード) 中に右側へ出す案内 + 解除ボタン。
const PanelArrangeBanner = ({ unresolvedCount, pageCount = 1, isFinalizing, onFinalize }) => (
  <div className="daiwari-panel-arrange-banner fixed right-48 top-1/2 z-[155] w-44 -translate-y-1/2 overflow-hidden rounded-2xl border border-slate-200/80 bg-white/90 p-2.5 shadow-[0_16px_40px_-16px_rgba(15,23,42,0.35)] backdrop-blur-xl animate-in fade-in slide-in-from-right-2 duration-200">
    <div className="flex items-center gap-2.5">
      <div className="daiwari-panel-arrange-icon relative flex h-8 w-8 flex-none items-center justify-center rounded-xl bg-slate-900 text-white shadow-sm">
        <ArrowLeftRight size={15} strokeWidth={2.25} />
        <span className="absolute -right-0.5 -top-0.5 h-2 w-2 rounded-full border-2 border-white bg-sky-400" />
      </div>
      <div className="min-w-0 flex-1">
        <p className="text-[11px] font-extrabold tracking-wide text-slate-800">ホバリング中</p>
        <p className="mt-0.5 truncate text-[9px] font-medium text-slate-500">
          {pageCount > 1 ? '押したまま2ページ間を移動' : '押したままコマへ移動'}
        </p>
      </div>
      <span className={`flex-none rounded-full px-1.5 py-0.5 text-[9px] font-extrabold ${unresolvedCount > 0
        ? 'bg-amber-100 text-amber-700'
        : 'bg-emerald-100 text-emerald-700'
        }`}>
        {unresolvedCount > 0 ? `残り${unresolvedCount}` : '完了'}
      </span>
    </div>

    <div className="mt-2 flex items-center gap-3 border-t border-slate-100 pt-2 text-[9px] font-bold text-slate-500">
      <span className="inline-flex items-center gap-1">
        <span className="h-2.5 w-2.5 rounded-[3px] border border-dashed border-sky-400 bg-sky-50" />
        浮遊中
      </span>
      <span className="inline-flex items-center gap-1">
        <span className="h-2.5 w-2.5 rounded-[3px] border border-emerald-400 bg-emerald-50" />
        割付済み
      </span>
    </div>

    <button
      type="button"
      onClick={onFinalize}
      disabled={isFinalizing}
      className={`mt-2 flex h-8 w-full items-center justify-center gap-1.5 rounded-xl text-[10px] font-extrabold transition-all ${unresolvedCount > 0
        ? 'bg-slate-100 text-slate-600 hover:bg-slate-200 hover:text-slate-800'
        : 'bg-slate-900 text-white shadow-sm hover:bg-slate-700 hover:shadow-md'
        } focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-sky-400 focus-visible:ring-offset-2 disabled:cursor-wait disabled:opacity-60`}
      title={unresolvedCount > 0 ? '未配置画像をすべて割り付けてください' : '配置を保存してホバリングを解除'}
    >
      {isFinalizing ? <Loader2 size={13} className="animate-spin" /> : <Check size={13} />}
      {isFinalizing ? '保存中…' : 'ホバリングを解除'}
    </button>
  </div>
);

export default PanelArrangeBanner;
