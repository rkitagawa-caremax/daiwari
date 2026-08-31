import React from 'react';
import { ChevronDown, ChevronUp } from 'lucide-react';

// 上部の操作バー (ヘッダー + コンテンツ内コントロール) の表示/非表示トグル。右上に固定表示。
const TopBarsToggleButton = ({ isTopBarsVisible, onToggle }) => (
  <button
    type="button"
    onClick={onToggle}
    className={`fixed right-3 z-[90] flex h-9 w-9 items-center justify-center rounded-xl border border-slate-200 bg-white/90 text-slate-500 shadow-md backdrop-blur transition-all duration-300 hover:bg-white hover:text-slate-700 hover:shadow-lg ${isTopBarsVisible ? 'top-[8rem]' : 'top-2'}`}
    title={isTopBarsVisible ? '上部の操作バーを隠す' : '上部の操作バーを表示'}
    aria-label={isTopBarsVisible ? '上部の操作バーを隠す' : '上部の操作バーを表示'}
    aria-pressed={!isTopBarsVisible}
  >
    {isTopBarsVisible ? <ChevronUp size={17} /> : <ChevronDown size={17} />}
  </button>
);

export default TopBarsToggleButton;
