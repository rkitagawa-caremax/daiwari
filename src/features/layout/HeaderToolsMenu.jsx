import {
  AlertCircle,
  CheckSquare,
  ChevronDown,
  FileSpreadsheet,
  Lightbulb,
  Lock,
  Plus,
  Sparkles,
  Tag,
  Unlock,
  Wrench
} from 'lucide-react';

const EdgeAiLaunchButton = ({ onLaunch }) => (
  <div className="group/edge-ai relative flex h-12 w-[92px] flex-shrink-0 items-center justify-center">
    <button
      type="button"
      onClick={() => onLaunch('search')}
      className="daiwari-ai-launch group relative flex h-12 w-full items-center justify-center px-2"
      title="AIアシストを開く"
      aria-label="AIアシストを開く"
      aria-haspopup="menu"
    >
      <span className="relative z-[2] flex h-[46px] items-center justify-center">
        <img src="/daiwari-kun.png" alt="" draggable="false" className="daiwari-ai-launch-image daiwari-ai-launch-frame daiwari-ai-launch-frame-front h-[46px] w-auto select-none object-contain" />
        <img src="/daiwari-kun-left.png" alt="" draggable="false" className="daiwari-ai-launch-image daiwari-ai-launch-frame daiwari-ai-launch-frame-left absolute inset-0 m-auto h-[46px] w-auto select-none object-contain" />
        <img src="/daiwari-kun-closed.png" alt="" draggable="false" className="daiwari-ai-launch-image daiwari-ai-launch-frame daiwari-ai-launch-frame-closed absolute inset-0 m-auto h-[46px] w-auto select-none object-contain" />
      </span>
      <span aria-hidden="true" className="daiwari-ai-sparkle daiwari-ai-sparkle-one">✦</span>
      <span aria-hidden="true" className="daiwari-ai-sparkle daiwari-ai-sparkle-two">✦</span>
      <span aria-hidden="true" className="daiwari-ai-sparkle daiwari-ai-sparkle-three">✦</span>
    </button>
    <div className="pointer-events-none absolute left-1/2 top-full z-[110] w-[316px] -translate-x-1/2 -translate-y-1 pt-2 opacity-0 transition duration-200 group-hover/edge-ai:pointer-events-auto group-hover/edge-ai:translate-y-0 group-hover/edge-ai:opacity-100 group-focus-within/edge-ai:pointer-events-auto group-focus-within/edge-ai:translate-y-0 group-focus-within/edge-ai:opacity-100">
      <div className="relative grid grid-cols-2 gap-2 overflow-hidden rounded-[22px] border border-white/90 bg-gradient-to-br from-white via-violet-50/95 to-cyan-50/90 p-2 shadow-[0_16px_40px_rgba(76,68,151,0.22)] ring-1 ring-violet-100/70 backdrop-blur-xl" role="menu" aria-label="AIアシストのショートカット">
        <span aria-hidden="true" className="absolute -right-8 -top-10 h-24 w-24 rounded-full bg-cyan-200/30 blur-2xl" />
        <span aria-hidden="true" className="absolute -bottom-12 -left-6 h-24 w-24 rounded-full bg-fuchsia-200/30 blur-2xl" />
        <button
          type="button"
          role="menuitem"
          onClick={() => onLaunch('advisor')}
          className="group/shortcut relative flex min-w-0 items-center justify-center gap-2 overflow-hidden rounded-2xl border border-violet-100 bg-white/90 px-3 py-3 text-xs font-black shadow-[0_5px_16px_rgba(124,92,214,0.10)] transition duration-200 hover:-translate-y-0.5 hover:border-violet-200 hover:shadow-[0_9px_22px_rgba(124,92,214,0.19)] focus:-translate-y-0.5 focus:border-violet-300 focus:outline-none focus:ring-2 focus:ring-violet-200"
        >
          <span aria-hidden="true" className="absolute inset-0 bg-gradient-to-r from-violet-50/0 via-fuchsia-50/80 to-cyan-50/0 opacity-0 transition group-hover/shortcut:opacity-100" />
          <span className="relative flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-violet-500 to-fuchsia-400 text-white shadow-[0_4px_10px_rgba(139,92,246,0.28)]"><Lightbulb size={16} /></span>
          <span className="daiwari-ai-shortcut-label daiwari-ai-shortcut-label-violet relative whitespace-nowrap">台割診断</span>
          <span aria-hidden="true" className="daiwari-ai-shortcut-sparkle daiwari-ai-shortcut-sparkle-one">✦</span>
          <span aria-hidden="true" className="daiwari-ai-shortcut-sparkle daiwari-ai-shortcut-sparkle-two">✦</span>
        </button>
        <button
          type="button"
          role="menuitem"
          onClick={() => onLaunch('search')}
          className="group/shortcut relative flex min-w-0 items-center justify-center gap-2 overflow-hidden rounded-2xl border border-cyan-100 bg-white/90 px-3 py-3 text-xs font-black shadow-[0_5px_16px_rgba(14,165,233,0.10)] transition duration-200 hover:-translate-y-0.5 hover:border-cyan-200 hover:shadow-[0_9px_22px_rgba(14,165,233,0.18)] focus:-translate-y-0.5 focus:border-cyan-300 focus:outline-none focus:ring-2 focus:ring-cyan-200"
        >
          <span aria-hidden="true" className="absolute inset-0 bg-gradient-to-r from-cyan-50/0 via-sky-50/90 to-violet-50/0 opacity-0 transition group-hover/shortcut:opacity-100" />
          <span className="relative flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-sky-400 to-cyan-400 text-white shadow-[0_4px_10px_rgba(14,165,233,0.25)]"><Sparkles size={16} /></span>
          <span className="daiwari-ai-shortcut-label daiwari-ai-shortcut-label-blue relative whitespace-nowrap">商品提案</span>
          <span aria-hidden="true" className="daiwari-ai-shortcut-sparkle daiwari-ai-shortcut-sparkle-one">✦</span>
          <span aria-hidden="true" className="daiwari-ai-shortcut-sparkle daiwari-ai-shortcut-sparkle-two">✦</span>
        </button>
      </div>
    </div>
  </div>
);

// ヘッダー右端の「ツール」ボタンと、そのポップアップメニュー
// (画面ロック / 選択モード / Q / ラベル強調 / 空き強調 / ページ追加 / 出力)。
// ポップアップはヘッダー (position: relative) を基準に absolute 配置するため、ヘッダー直下に置く。
const HeaderToolsMenu = ({
  isOpen,
  onToggle,
  onClose,
  viewMode,
  isLocked,
  isPageSelectionMode,
  isQuickHelpMode,
  highlightLabels,
  highlightEmpty,
  lockHoldFiredRef,
  onStartLockHold,
  onCancelLockHold,
  onTogglePageSelectionMode,
  onToggleQuickHelpMode,
  onToggleHighlightLabels,
  onToggleHighlightEmpty,
  onAddSheet,
  onOpenEdgeAi,
  onExportCSV,
  onShowQuickHelp,
  onHideQuickHelp
}) => (
  <>
    <div className="flex items-center gap-2 flex-shrink-0 ml-4">
      <EdgeAiLaunchButton
        onLaunch={(initialTab) => {
          onClose();
          onOpenEdgeAi(initialTab);
        }}
      />

      <button
        type="button"
        onClick={onToggle}
        onMouseEnter={(e) => onShowQuickHelp(e, 'ツール', '画面ロック・選択モード・強調表示・CSV出力などの機能をまとめています。')}
        onMouseLeave={onHideQuickHelp}
        className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full transition-all duration-300 border whitespace-nowrap ${isOpen ? 'bg-slate-700 text-white border-slate-700 shadow-md' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
        title="ツール"
        aria-expanded={isOpen}
        aria-haspopup="true"
      >
        <Wrench size={14} strokeWidth={2.5} /> <span>ツール</span>
        <ChevronDown size={13} className={`transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`} />
      </button>
    </div>

    {isOpen && (
      <>
        <div className="fixed inset-0 z-[94]" onClick={onClose} aria-hidden="true" />
        <div
          className="absolute right-6 top-full z-[95] mt-2 flex items-center gap-2 rounded-2xl border border-slate-200 bg-white/95 p-3 shadow-xl backdrop-blur-md animate-in fade-in slide-in-from-top-2 duration-200"
          role="menu"
          aria-label="ツールメニュー"
        >
          {/* 画面ロックボタン: 2秒長押しでトグル。ロック中は編集系を一律 no-op、閲覧・画面切替・ページ移動は可能。 */}
          <button
            type="button"
            onPointerDown={onStartLockHold}
            onPointerUp={onCancelLockHold}
            onPointerLeave={onCancelLockHold}
            onPointerCancel={onCancelLockHold}
            onClick={(e) => {
              // 長押し未満の単発クリックでは何もしない (誤発動防止)。
              if (!lockHoldFiredRef.current) {
                e.preventDefault();
              }
              lockHoldFiredRef.current = false;
            }}
            onMouseEnter={(e) => onShowQuickHelp(e, isLocked ? '画面ロック中' : '画面ロック', isLocked ? '2秒長押しで解除します。閲覧・画面切替・ページ移動は引き続き使えます。' : '鍵を2秒長押しで編集を一時停止します。閲覧・画面切替・ページ移動は引き続き可能です。')}
            onMouseLeave={onHideQuickHelp}
            title={isLocked ? '画面ロック中 (2秒長押しで解除)' : '画面をロック (2秒長押し)'}
            className={`flex items-center justify-center w-10 h-10 rounded-full transition-colors ${
              isLocked
                ? 'bg-rose-100 text-rose-600 border-2 border-rose-300 shadow-inner'
                : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-50'
            }`}
            style={{ touchAction: 'none' }}
          >
            {isLocked ? <Lock size={18} /> : <Unlock size={18} />}
          </button>

          {viewMode === 'overview' && (
            <button
              onClick={() => {
                onTogglePageSelectionMode();
                onClose();
              }}
              onMouseEnter={(e) => onShowQuickHelp(e, '選択モード', '複数ページを選択して、入れ替え・画像解除・削除を行います。')}
              onMouseLeave={onHideQuickHelp}
              className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full transition-all duration-300 border whitespace-nowrap ${isPageSelectionMode ? 'bg-indigo-600 text-white border-indigo-600 shadow-md' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              title="複数ページを選択して削除"
            >
              <CheckSquare size={14} strokeWidth={2.5} /> <span>選択モード</span>
            </button>
          )}

          <button
            onClick={onToggleQuickHelpMode}
            className={`w-9 h-9 flex-shrink-0 rounded-full border text-[14px] font-extrabold leading-none transition-all ${isQuickHelpMode
              ? 'bg-sky-600 text-white border-sky-500 ring-4 ring-sky-300/50 shadow-[0_0_20px_rgba(56,189,248,0.55)]'
              : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}
            title="クイックヘルプ"
          >
            Q
          </button>

          {!isPageSelectionMode && viewMode === 'overview' && (
            <button
              onClick={onToggleHighlightLabels}
              onMouseEnter={(e) => onShowQuickHelp(e, 'ラベル強調', 'ラベルが1つ以上あるコマを緑色で強調表示します。もう一度押すと解除します。')}
              onMouseLeave={onHideQuickHelp}
              className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full transition-all duration-300 border whitespace-nowrap ${highlightLabels ? 'bg-emerald-500 text-white border-emerald-600 shadow-md' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              title="自由ラベルがあるコマを緑色で強調表示"
            >
              <Tag size={14} strokeWidth={2.5} /> <span>ラベル強調</span>
            </button>
          )}

          <button
            onClick={onToggleHighlightEmpty}
            onMouseEnter={(e) => onShowQuickHelp(e, '空き強調', '空きコマを赤色で強調表示します。全体表示時の確認に使います。')}
            onMouseLeave={onHideQuickHelp}
            className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full transition-all duration-300 border whitespace-nowrap ${highlightEmpty ? 'bg-rose-500 text-white border-rose-600 shadow-md' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
            title="空きコマを赤色で強調表示"
          >
            <AlertCircle size={14} strokeWidth={2.5} /> <span>空き強調</span>
          </button>

          <button
            onClick={onAddSheet}
            onMouseEnter={(e) => onShowQuickHelp(e, '+ページ追加', '新しいページを末尾に追加します。続けてクリックすると複数追加できます。')}
            onMouseLeave={onHideQuickHelp}
            className="flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full border border-indigo-600 bg-indigo-600 text-white shadow-sm hover:bg-indigo-700 transition-all duration-300 active:scale-95 whitespace-nowrap"
            title="新しいページを末尾に追加"
          >
            <Plus size={14} strokeWidth={3} /> ページ追加
          </button>

          <button
            onClick={() => {
              onClose();
              onExportCSV();
            }}
            onMouseEnter={(e) => onShowQuickHelp(e, '出力', '現在のページ情報をCSVで出力します。外部共有やバックアップに使えます。')}
            onMouseLeave={onHideQuickHelp}
            className="flex items-center gap-2 px-5 py-2.5 text-sm font-bold rounded-full transition-all duration-300 border-2 border-emerald-500 bg-white text-emerald-600 hover:bg-emerald-50 shadow-sm hover:shadow whitespace-nowrap"
            title="ページ情報をCSVでダウンロード"
          >
            <FileSpreadsheet size={16} strokeWidth={2.5} /> <span>出力</span>
          </button>
        </div>
      </>
    )}
  </>
);

export default HeaderToolsMenu;
