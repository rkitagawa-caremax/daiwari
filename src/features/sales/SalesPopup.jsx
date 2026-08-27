import React, { useEffect, useRef, useState } from 'react';

const SalesPopup = React.memo(({ data, position, onMouseEnter, onMouseLeave }) => {
  const popupRef = useRef(null);
  const [offset, setOffset] = useState({ x: 10, y: 10 });

  useEffect(() => {
    if (popupRef.current && position) {
      const rect = popupRef.current.getBoundingClientRect();
      const viewportWidth = window.innerWidth;
      const viewportHeight = window.innerHeight;

      let nextX = 10;
      let nextY = 10;

      if (position.x + rect.width + 20 > viewportWidth) {
        nextX = -rect.width - 10;
      }
      if (position.y + rect.height + 20 > viewportHeight) {
        nextY = viewportHeight - (position.y + rect.height + 20);
      }
      setOffset({ x: nextX, y: nextY });
    }
  }, [position, data]);

  if (!data || !position) return null;

  const totalCount = data.reduce((sum, item) => sum + (parseInt(item.count) || 0), 0);

  return (
    <div
      ref={popupRef}
      className="fixed z-[100] bg-white/95 backdrop-blur-md rounded-2xl shadow-2xl border border-slate-200 p-4 w-72 animate-in fade-in zoom-in-95 duration-200 pointer-events-auto"
      style={{
        top: position.y + offset.y,
        left: position.x + offset.x,
        boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1), 0 0 0 1px rgb(0 0 0 / 0.05)'
      }}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
    >
      <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-3">
        <div className="flex flex-col">
          <span className="text-[10px] uppercase tracking-wider font-bold text-slate-400">Sales Record</span>
          <span className="text-xs font-bold text-slate-600">販売実績詳細</span>
        </div>
        <div className="flex flex-col items-end">
          <span className="text-2xl font-black text-indigo-600 font-mono leading-none">{totalCount.toLocaleString()}</span>
          <span className="text-[10px] text-slate-400 font-bold uppercase">Total Units</span>
        </div>
      </div>
      <ul className="space-y-2 max-h-64 overflow-y-auto pr-1 custom-scrollbar">
        {data.map((item, index) => (
          <li key={index} className="text-[11px] leading-tight flex justify-between gap-3 py-2 border-b border-slate-50 last:border-0 hover:bg-slate-50/50 transition-colors rounded-lg px-1">
            <div className="flex flex-col gap-0.5 min-w-0">
              <span className="font-bold text-slate-800 truncate">{item.name}</span>
              <span className="text-slate-500 truncate text-[9px] opacity-80">{item.spec}</span>
            </div>
            <span className="font-mono font-bold text-indigo-500 bg-indigo-50 px-1.5 py-0.5 rounded ml-auto self-center">{item.count}</span>
          </li>
        ))}
      </ul>
    </div>
  );
});

export default SalesPopup;
