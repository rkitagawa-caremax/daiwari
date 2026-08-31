import React from 'react';
import { BarChart2, Grid, List, LogOut, Redo2, Undo2 } from 'lucide-react';

// 画面最上部のナビゲーションバー (M3 Expressive Style)。
// 左: ロゴ / ログイン情報 / 戻る・進む / 詳細・全体 切替 / (選択モード時) 一括操作 / (詳細時) 実績モード
// 右: ツールメニュー。selectionToolbar と toolsMenu は App 側で組み立てた要素をスロットとして受け取る
// (toolsMenu のポップアップは このヘッダーの position: relative を基準に absolute 配置される)。
const AppHeader = ({
  isLocalMode,
  signedInUserName,
  onLogoTap,
  onLogout,
  onUndo,
  onRedo,
  viewMode,
  onSelectViewMode,
  isPageSelectionMode,
  isSalesMode,
  isSalesLookupOpen,
  onSalesModeClick,
  onSalesModeLongPressStart,
  onSalesModeLongPressEnd,
  onShowQuickHelp,
  onHideQuickHelp,
  selectionToolbar,
  toolsMenu
}) => (
  <div className="h-14 flex items-center justify-between px-4 z-30 flex-shrink-0 relative transition-all" style={{ background: 'var(--m3-surface)', color: 'var(--m3-on-surface)' }}>
    <div className="flex items-center gap-5 flex-shrink-0">
      <div className="flex items-center">
        <div
          className="p-0.5 bg-white shadow-sm cursor-pointer select-none"
          style={{ borderRadius: 'var(--m3-shape-corner-md)' }}
          onClick={onLogoTap}
          title="台"
        >
          <img
            src="/logo.jpg"
            alt="台割君"
            className="h-10 w-10 object-contain transition-transform hover:scale-105"
            style={{ borderRadius: 'calc(var(--m3-shape-corner-md) - 2px)' }}
          />
        </div>
      </div>

      <div className="h-6 w-px mx-1 opacity-50" style={{ background: 'var(--m3-outline-variant)' }}></div>

      {!isLocalMode && signedInUserName && (
        <div className="flex items-center gap-1 text-slate-400">
          <span className="max-w-[9rem] truncate text-[11px] font-medium text-slate-500" title={signedInUserName}>{signedInUserName}</span>
          <button
            type="button"
            onClick={onLogout}
            className="flex h-6 w-6 items-center justify-center rounded-full text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-600"
            title="ログアウト"
            aria-label="ログアウト"
          >
            <LogOut size={13} />
          </button>
        </div>
      )}

      {/* 戻る / 進む (アカウント単位の undo / redo) */}
      <div className="flex items-center gap-1 mr-1">
        <button
          type="button"
          onClick={onUndo}
          onMouseEnter={(e) => onShowQuickHelp(e, '戻る', '直前の編集操作を取り消します (Ctrl+Z)。')}
          onMouseLeave={onHideQuickHelp}
          title="戻る (Ctrl+Z)"
          aria-label="直前の編集操作を戻す"
          className="flex items-center justify-center w-8 h-8 rounded-full border border-slate-200 bg-white text-slate-600 transition-colors hover:bg-slate-50"
        >
          <Undo2 size={16} />
        </button>
        <button
          type="button"
          onClick={onRedo}
          onMouseEnter={(e) => onShowQuickHelp(e, '進む', '戻した操作をやり直します (Ctrl+Y / Ctrl+Shift+Z)。')}
          onMouseLeave={onHideQuickHelp}
          title="進む (Ctrl+Y)"
          aria-label="戻した編集操作をやり直す"
          className="flex items-center justify-center w-8 h-8 rounded-full border border-slate-200 bg-white text-slate-600 transition-colors hover:bg-slate-50"
        >
          <Redo2 size={16} />
        </button>
      </div>

      <div className="flex p-1 rounded-full transition-all" style={{ border: '1px solid var(--m3-outline)', background: 'var(--m3-surface)' }}>
        <button
          onClick={() => onSelectViewMode('list')}
          onMouseEnter={(e) => onShowQuickHelp(e, '詳細', 'ページ単位で編集する表示に切り替えます。')}
          onMouseLeave={onHideQuickHelp}
          className={`flex items-center gap-1.5 px-4 py-1.5 text-sm font-medium rounded-full transition-all duration-300 whitespace-nowrap`}
          style={viewMode === 'list' || viewMode === 'single' ? { background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
        >
          <List size={16} /> <span className="hidden sm:inline">詳細</span>
        </button>
        <button
          onClick={() => onSelectViewMode('overview')}
          onMouseEnter={(e) => onShowQuickHelp(e, '全体', '全ページを一覧で表示します。コマの全体把握に使います。')}
          onMouseLeave={onHideQuickHelp}
          className={`flex items-center gap-1.5 px-4 py-1.5 text-sm font-medium rounded-full transition-all duration-300 whitespace-nowrap`}
          style={viewMode === 'overview' ? { background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
        >
          <Grid size={16} /> <span className="hidden sm:inline">全体</span>
        </button>
      </div>

      {selectionToolbar}

      {/* 実績モード Toggle - 詳細表示時のみ */}
      {!isPageSelectionMode && (viewMode === 'list' || viewMode === 'single') && (
        <button
          onClick={onSalesModeClick}
          onMouseDown={onSalesModeLongPressStart}
          onMouseUp={onSalesModeLongPressEnd}
          onMouseLeave={() => { onSalesModeLongPressEnd(); onHideQuickHelp(); }}
          onTouchStart={onSalesModeLongPressStart}
          onTouchEnd={onSalesModeLongPressEnd}
          onTouchCancel={onSalesModeLongPressEnd}
          onMouseEnter={(e) => onShowQuickHelp(e, '実績モード', 'クリックで重ね表示のON/OFF。2秒長押しで介援隊コード検索POPを開きます。')}
          className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border text-sm font-bold transition-all duration-300 ml-2 whitespace-nowrap
             ${isSalesLookupOpen
              ? 'bg-violet-500/15 border-violet-500 text-violet-700 shadow-[0_0_18px_rgba(139,92,246,0.55)] animate-pulse'
              : isSalesMode
              ? 'bg-emerald-500/10 border-emerald-500 text-emerald-400 shadow-[0_0_15px_rgba(16,185,129,0.3)]'
              : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}`}
          title="クリック: 実績モード切替 / 2秒長押し: コード実績検索"
        >
          <BarChart2 size={18} />
          <span className="hidden xl:inline">実績モード {isSalesMode ? 'ON' : 'OFF'}</span>
        </button>
      )}
    </div>

    {toolsMenu}
  </div>
);

export default AppHeader;
