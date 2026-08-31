import React from 'react';
import { ArrowLeftRight, Check, Loader2 } from 'lucide-react';

// 画像ホバリング (パネル配置モード) 中に右側へ出す案内 + 解除ボタン。
const PanelArrangeBanner = ({ unresolvedCount, pageCount = 1, isFinalizing, onFinalize }) => (
  <div className="daiwari-panel-arrange-banner fixed right-48 top-1/2 z-[155] w-48 -translate-y-1/2 rounded-2xl border border-sky-200 bg-white/95 p-3 shadow-xl backdrop-blur-md">
    <div className="flex items-center gap-2">
      <div className="daiwari-panel-arrange-icon flex h-8 w-8 flex-none items-center justify-center rounded-full bg-sky-100 text-sky-700">
        <ArrowLeftRight size={17} />
      </div>
      <div className="min-w-0">
        <p className="text-xs font-extrabold text-slate-800">画像ホバリング中</p>
        <p className={`text-[10px] font-bold ${unresolvedCount > 0 ? 'text-amber-600' : 'text-emerald-600'}`}>
          {unresolvedCount > 0 ? `未配置 ${unresolvedCount}件` : '解除できます'}
        </p>
      </div>
    </div>
    <p className="mt-2 text-[10px] leading-relaxed text-slate-500">
      {pageCount > 1
        ? '画像を押したまま2ページ間のコマへ移動できます。青の点線は浮遊中、黄緑の枠は割付済みです。'
        : '画像を押したままコマへ移動できます。青の点線は浮遊中、黄緑の枠は割付済みです。'}
    </p>
    <button
      type="button"
      onClick={onFinalize}
      disabled={isFinalizing}
      className={`mt-3 flex w-full items-center justify-center gap-1.5 rounded-xl border px-3 py-2 text-xs font-extrabold shadow-sm transition-all ${unresolvedCount > 0
        ? 'border-amber-300 bg-amber-50 text-amber-700 hover:bg-amber-100'
        : 'border-emerald-300 bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
        } disabled:cursor-wait disabled:opacity-60`}
      title={unresolvedCount > 0 ? '未配置画像をすべて割り付けてください' : '配置を保存してホバリングを解除'}
    >
      {isFinalizing ? <Loader2 size={14} className="animate-spin" /> : <Check size={14} />}
      ホバリングを解除
    </button>
  </div>
);

export default PanelArrangeBanner;
