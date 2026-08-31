import React from 'react';
import { ChevronLeft, ChevronRight } from 'lucide-react';

// コンテンツ領域の先頭に出す「ジャンル絞り込み」+「ページ移動」コントロール。
// data-* 属性は作業ログ計測やスタイル参照で使うため維持する。
const ContentHeaderControls = ({
  viewMode,
  genres,
  genreFilter,
  onChangeGenreFilter,
  currentIndex,
  totalCount,
  activeSheetId,
  isSalesMode,
  onNavigate
}) => (
  <div
    data-display-context-controls="true"
    className={`grid grid-cols-[minmax(0,1fr)_auto_minmax(0,1fr)] items-center gap-3 px-1 py-1 ${viewMode === 'single'
      ? 'sticky top-0 z-40 backdrop-blur-sm'
      : ''}`}
  >
    <div data-genre-filter-control="true" className="flex min-w-0 items-center gap-1.5 justify-self-start text-slate-400">
      <span className="text-[10px] font-medium tracking-wide">ジャンル</span>
      <select
        value={genreFilter}
        onChange={(e) => onChangeGenreFilter(e.target.value)}
        className="max-w-36 cursor-pointer border-none bg-transparent py-1 pr-1 text-xs font-medium text-slate-500 outline-none transition-colors hover:text-slate-700 focus:text-indigo-600"
        aria-label="表示ジャンル"
      >
        <option value="all">全て表示</option>
        {genres.map(g => (
          <option key={g.id} value={g.id}>{g.label}</option>
        ))}
      </select>
    </div>

    <nav data-page-navigation="true" className="flex items-center justify-center gap-1 text-slate-400" aria-label="ページ移動">
      <button
        type="button"
        onClick={() => onNavigate('prev')}
        disabled={viewMode !== 'single' || currentIndex <= 0}
        className="rounded-md p-1 text-slate-400 transition-colors hover:bg-slate-200/50 hover:text-slate-600 disabled:cursor-not-allowed disabled:opacity-20"
        aria-label="前のページ"
      >
        <ChevronLeft size={15} />
      </button>
      <div className="flex min-w-[6.25rem] items-baseline justify-center gap-1.5 whitespace-nowrap">
        <span className="text-[9px] font-medium uppercase tracking-wider text-slate-400">Page</span>
        <span className={`font-mono text-xs font-semibold ${isSalesMode ? 'text-slate-300' : 'text-slate-500'}`}>
          {viewMode === 'single' && activeSheetId ? currentIndex + 1 : '-'}
          <span className="mx-1 font-normal text-slate-300">/</span>
          <span className="font-medium text-slate-400">{totalCount}</span>
        </span>
      </div>
      <button
        type="button"
        onClick={() => onNavigate('next')}
        disabled={viewMode !== 'single' || currentIndex === -1 || currentIndex >= totalCount - 1}
        className="rounded-md p-1 text-slate-400 transition-colors hover:bg-slate-200/50 hover:text-slate-600 disabled:cursor-not-allowed disabled:opacity-20"
        aria-label="次のページ"
      >
        <ChevronRight size={15} />
      </button>
    </nav>

  </div>
);

export default ContentHeaderControls;
