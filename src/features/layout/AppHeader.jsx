import React from 'react';
import { BarChart2, ChartSpline, FileDiff, Grid, Hash, List, LogOut, PieChart, Redo2, Undo2 } from 'lucide-react';

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
  showPanelCodes,
  onTogglePanelCodes,
  isPageSelectionMode,
  isSalesMode,
  isSalesChartMode,
  salesDisplayMode = 'quantity',
  isSalesLookupOpen,
  salesPeriodOptions = [],
  activeSalesPeriod,
  salesPeriodMeta = {},
  isSalesPeriodLoading = false,
  onSelectSalesPeriod,
  onSalesModeClick,
  onSalesChartModeClick,
  onGrossProfitModeClick,
  onSalesModeLongPressStart,
  onSalesModeLongPressEnd,
  catalogChangeCount,
  catalogChangeFileName,
  isCatalogDiffMode,
  onCatalogDiffModeClick,
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

      {/* 介援隊コード表示の切替 — 詳細・全体どちらでも常設 */}
      <button
        type="button"
        onClick={onTogglePanelCodes}
        onMouseEnter={(event) => onShowQuickHelp(event, '介援隊コード表示', '各コマ右上に表示している介援隊コードの表示/非表示を切り替えます。')}
        onMouseLeave={onHideQuickHelp}
        className={`ml-1 flex items-center gap-2 whitespace-nowrap rounded-xl border px-3 py-1.5 text-sm font-bold transition-all duration-300 ${showPanelCodes
          ? 'border-slate-200 bg-white text-slate-600 hover:border-slate-300 hover:bg-slate-50'
          : 'border-slate-500 bg-slate-600 text-white shadow-md'}`}
        aria-pressed={!showPanelCodes}
        title={showPanelCodes ? '介援隊コードを非表示にする' : '介援隊コードを表示する'}
      >
        <Hash size={16} strokeWidth={2.6} />
        <span className="hidden xl:inline">コード {showPanelCodes ? 'ON' : 'OFF'}</span>
      </button>

      {selectionToolbar}

      {/* 実績モード Toggle - 詳細表示時のみ */}
      {!isPageSelectionMode && (viewMode === 'list' || viewMode === 'single') && (
        <>
          <button
            onClick={onSalesModeClick}
            onMouseDown={onSalesModeLongPressStart}
            onMouseUp={onSalesModeLongPressEnd}
            onMouseLeave={() => { onSalesModeLongPressEnd(); onHideQuickHelp(); }}
            onTouchStart={onSalesModeLongPressStart}
            onTouchEnd={onSalesModeLongPressEnd}
            onTouchCancel={onSalesModeLongPressEnd}
            onMouseEnter={(e) => onShowQuickHelp(e, '実績モード', 'クリックで重ね表示のON/OFF。2秒長押しで介援隊コード検索POPを開きます。')}
            className={`ml-2 flex items-center gap-2 whitespace-nowrap rounded-xl border px-3 py-1.5 text-sm font-bold transition-all duration-300
              ${isSalesLookupOpen
                ? 'animate-pulse border-violet-500 bg-violet-500/15 text-violet-700 shadow-[0_0_18px_rgba(139,92,246,0.55)]'
                : isSalesMode && salesDisplayMode === 'quantity' && !isSalesChartMode
                  ? 'border-emerald-500 bg-emerald-500/10 text-emerald-400 shadow-[0_0_15px_rgba(16,185,129,0.3)]'
                  : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'}`}
            title="クリック: 実績モード切替 / 2秒長押し: コード実績検索"
          >
            <BarChart2 size={18} />
            <span className="hidden xl:inline">実績モード {isSalesMode && salesDisplayMode === 'quantity' && !isSalesChartMode ? 'ON' : 'OFF'}</span>
          </button>

          <button
            type="button"
            onClick={onSalesChartModeClick}
            onMouseEnter={(event) => onShowQuickHelp(event, '月別グラフ', '表示中の詳細ページにある全商品の月別売上グラフを一括で表示します。')}
            onMouseLeave={onHideQuickHelp}
            className={`ml-1 flex items-center gap-2 whitespace-nowrap rounded-xl border px-3 py-1.5 text-sm font-bold transition-all duration-300 ${isSalesChartMode
              ? 'border-blue-500 bg-blue-500/10 text-blue-600 shadow-[0_0_15px_rgba(59,130,246,0.24)]'
              : 'border-slate-200 bg-white text-blue-500 hover:border-blue-300 hover:bg-blue-50'}`}
            aria-pressed={isSalesChartMode}
            title={isSalesChartMode ? '月別グラフの一括表示を解除' : '月別グラフを全コマに一括表示'}
          >
            <ChartSpline size={18} strokeWidth={2.4} />
            <span className="hidden xl:inline">月別グラフ {isSalesChartMode ? 'ON' : 'OFF'}</span>
          </button>

          <button
            type="button"
            onClick={onGrossProfitModeClick}
            onMouseEnter={(event) => onShowQuickHelp(event, '粗利データ', 'コマごとの粗利率を円グラフ、粗利総額を金額で表示します。')}
            onMouseLeave={onHideQuickHelp}
            className={`ml-1 flex items-center gap-2 whitespace-nowrap rounded-xl border px-3 py-1.5 text-sm font-bold transition-all duration-300 ${isSalesMode && salesDisplayMode === 'grossProfit'
              ? 'border-yellow-400 bg-yellow-300/20 text-amber-700 shadow-[0_0_18px_rgba(250,204,21,0.55)]'
              : 'border-slate-200 bg-white text-amber-600 hover:border-yellow-300 hover:bg-yellow-50'}`}
            aria-pressed={isSalesMode && salesDisplayMode === 'grossProfit'}
            title="コマ上に粗利率と粗利総額を表示"
          >
            <PieChart size={18} strokeWidth={2.5} />
            <span className="hidden xl:inline">粗利データ</span>
          </button>

          {isSalesMode && salesPeriodOptions.length > 1 && (
            <div className="ml-1 flex items-center gap-0.5 rounded-xl border border-slate-200 bg-white p-0.5" role="group" aria-label="売上データの対象期間">
              {salesPeriodOptions.map((period) => {
                const isActive = period.id === activeSalesPeriod;
                const periodMeta = salesPeriodMeta?.[period.id];
                const hasData = salesDisplayMode === 'grossProfit'
                  ? (periodMeta?.metrics?.grossProfitAmount?.codes || 0) > 0
                  : !!periodMeta;
                return (
                  <button
                    key={period.id}
                    type="button"
                    onClick={() => onSelectSalesPeriod?.(period.id)}
                    onMouseEnter={(event) => onShowQuickHelp(event, period.label, hasData
                      ? `${period.description}の売上データを表示します。`
                      : `${period.description}の売上データはまだ取り込まれていません。設定から取り込めます。`)}
                    onMouseLeave={onHideQuickHelp}
                    className={`rounded-lg px-2.5 py-1 text-xs font-bold transition-colors ${isActive
                      ? `${salesDisplayMode === 'grossProfit' ? 'bg-yellow-400 text-amber-950 shadow-[0_0_12px_rgba(250,204,21,0.5)]' : isSalesChartMode ? 'bg-cyan-600 text-white' : 'bg-emerald-500 text-white'} shadow-sm`
                      : hasData ? 'text-slate-600 hover:bg-slate-100' : 'text-slate-300'}`}
                    aria-pressed={isActive}
                    title={hasData ? `${period.label}の売上を表示` : `${period.label}のデータは未取り込み`}
                  >
                    {period.label}
                  </button>
                );
              })}
              {isSalesPeriodLoading && <span className="px-1 text-[10px] font-bold text-slate-400">読込中…</span>}
            </div>
          )}

          {catalogChangeCount > 0 && (
            <button
              type="button"
              onClick={onCatalogDiffModeClick}
              onMouseEnter={(event) => onShowQuickHelp(event, '差分モード', `${catalogChangeFileName || '取込データ'}との変更 ${catalogChangeCount}件をコマ上に重ねて表示します。`)}
              onMouseLeave={onHideQuickHelp}
              className={`ml-1 flex items-center gap-2 whitespace-nowrap rounded-xl border px-3 py-1.5 text-sm font-bold transition-all duration-300 ${isCatalogDiffMode
                ? 'border-amber-500 bg-amber-500/10 text-amber-700 shadow-[0_0_15px_rgba(245,158,11,0.24)]'
                : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'}`}
              title="商品情報の変更をコマ上に表示"
            >
              <FileDiff size={17} />
              <span className="hidden xl:inline">差分 {catalogChangeCount}件 {isCatalogDiffMode ? 'ON' : 'OFF'}</span>
            </button>
          )}
        </>
      )}
    </div>

    {toolsMenu}
  </div>
);

export default AppHeader;
