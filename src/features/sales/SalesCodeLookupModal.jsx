import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { CheckSquare, GripVertical, Search, Square, X } from 'lucide-react';

import { normalizeCode } from '../../domain/productCodes';
import { clamp } from '../../lib/math';

const SalesCodeLookupModal = React.memo(({ isOpen, onClose, salesData, visibleCodes = null }) => {
  const [query, setQuery] = useState('');
  const [selectedCodes, setSelectedCodes] = useState([]);
  const [position, setPosition] = useState({ x: 24, y: 96 });
  const inputRef = useRef(null);
  const popupRef = useRef(null);
  const dragStateRef = useRef({ active: false, offsetX: 0, offsetY: 0 });
  const isScopedToPage = Array.isArray(visibleCodes);
  const visibleCodeSet = useMemo(() => {
    if (!isScopedToPage) return null;
    return new Set(
      visibleCodes
        .map((code) => normalizeCode(code))
        .filter(Boolean)
    );
  }, [isScopedToPage, visibleCodes]);

  const clampPopupPosition = useCallback((x, y) => {
    const rect = popupRef.current?.getBoundingClientRect();
    const width = rect?.width || 360;
    const height = rect?.height || Math.min(window.innerHeight * 0.78, 760);
    const maxX = Math.max(8, window.innerWidth - width - 8);
    const maxY = Math.max(8, window.innerHeight - height - 8);
    return {
      x: clamp(x, 8, maxX),
      y: clamp(y, 8, maxY)
    };
  }, []);

  useEffect(() => {
    if (!isOpen) return;
    setQuery('');
    setSelectedCodes([]);

    const nextX = Math.max(12, window.innerWidth - 392);
    const nextY = 92;
    setPosition(clampPopupPosition(nextX, nextY));

    setTimeout(() => {
      inputRef.current?.focus();
    }, 0);
  }, [isOpen, clampPopupPosition]);

  useEffect(() => {
    if (!isOpen) return;
    const handleResize = () => {
      setPosition((previous) => clampPopupPosition(previous.x, previous.y));
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, [isOpen, clampPopupPosition]);

  useEffect(() => {
    if (!isOpen) return;
    const handleEsc = (event) => {
      if (event.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleEsc);
    return () => window.removeEventListener('keydown', handleEsc);
  }, [isOpen, onClose]);

  const startDragging = useCallback((event) => {
    if (event?.button !== undefined && event.button !== 0) return;
    const rect = popupRef.current?.getBoundingClientRect();
    if (!rect) return;
    event.preventDefault();

    dragStateRef.current = {
      active: true,
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top
    };

    const handleMove = (moveEvent) => {
      if (!dragStateRef.current.active) return;
      const nextX = moveEvent.clientX - dragStateRef.current.offsetX;
      const nextY = moveEvent.clientY - dragStateRef.current.offsetY;
      setPosition(clampPopupPosition(nextX, nextY));
    };

    const handleUp = () => {
      dragStateRef.current.active = false;
      window.removeEventListener('pointermove', handleMove);
    };

    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp, { once: true });
  }, [clampPopupPosition]);

  const allEntries = useMemo(() => {
    if (!salesData || typeof salesData !== 'object') return [];
    return Object.entries(salesData).map(([code, items]) => {
      const normalizedItems = Array.isArray(items) ? items : [];
      const total = normalizedItems.reduce((sum, item) => sum + (parseInt(item.count) || 0), 0);
      return {
        code,
        items: normalizedItems,
        total
      };
    }).sort((left, right) => {
      if (right.total !== left.total) return right.total - left.total;
      return left.code.localeCompare(right.code);
    });
  }, [salesData]);

  const scopedEntries = useMemo(() => {
    if (!visibleCodeSet) return allEntries;
    return allEntries.filter((entry) => visibleCodeSet.has(normalizeCode(entry.code)));
  }, [allEntries, visibleCodeSet]);

  const normalizedQuery = normalizeCode(query);

  const filteredEntries = useMemo(() => {
    if (!normalizedQuery) return scopedEntries.slice(0, 250);
    return allEntries.filter((entry) => entry.code.includes(normalizedQuery)).slice(0, 250);
  }, [scopedEntries, allEntries, normalizedQuery]);

  useEffect(() => {
    if (!isOpen) return;
    setSelectedCodes((previous) => {
      const availableCodes = new Set(allEntries.map((entry) => entry.code));
      const kept = previous.filter((code) => availableCodes.has(code));
      const same = kept.length === previous.length && kept.every((code, index) => code === previous[index]);
      return same ? previous : kept;
    });
  }, [isOpen, allEntries]);

  const selectedCodeSet = useMemo(() => new Set(selectedCodes), [selectedCodes]);

  const selectedEntries = useMemo(() => {
    if (selectedCodes.length === 0) return [];
    return allEntries.filter((entry) => selectedCodeSet.has(entry.code));
  }, [allEntries, selectedCodeSet, selectedCodes.length]);

  const toggleSelectedCode = useCallback((code) => {
    setSelectedCodes((previous) => {
      if (previous.includes(code)) return previous.filter((item) => item !== code);
      return [...previous, code];
    });
  }, []);

  const selectAllFilteredCodes = useCallback(() => {
    setSelectedCodes((previous) => {
      const merged = new Set(previous);
      filteredEntries.forEach((entry) => merged.add(entry.code));
      return Array.from(merged);
    });
  }, [filteredEntries]);

  const clearSelectedCodes = useCallback(() => {
    setSelectedCodes([]);
  }, []);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[130] pointer-events-none">
      <div
        ref={popupRef}
        className="absolute pointer-events-auto w-[360px] max-w-[92vw] h-[78vh] max-h-[86vh] min-h-[520px] rounded-xl border border-slate-200 bg-white shadow-2xl overflow-hidden flex flex-col"
        style={{ left: position.x, top: position.y }}
      >
        <div className="flex items-center justify-between gap-2 px-3 py-2 border-b border-slate-200 bg-slate-50">
          <div
            className="flex items-center gap-1.5 cursor-move select-none text-slate-500"
            onPointerDown={startDragging}
            title="ドラッグして移動"
          >
            <GripVertical size={14} />
            <div>
              <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase leading-none">Sales Lookup</p>
              <p className="text-xs font-bold text-slate-700 leading-tight mt-0.5">介援隊コード別 実績確認</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-md text-slate-500 hover:bg-slate-200 transition-colors"
            title="閉じる"
          >
            <X size={15} />
          </button>
        </div>

        <div className="px-3 py-2 border-b border-slate-100">
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              ref={inputRef}
              type="text"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="介援隊コードを入力（例: A1234）"
              className="w-full pl-8 pr-2.5 py-1.5 rounded-md border border-slate-200 bg-white text-xs font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
          </div>
          <p className="mt-1.5 text-[10px] text-slate-500">
            該当コード <span className="font-bold text-slate-700">{filteredEntries.length}</span> 件
            {normalizedQuery ? ` / ${normalizedQuery}` : ''}
          </p>
          <div className="mt-1.5 flex items-center gap-2">
            <button
              onClick={selectAllFilteredCodes}
              className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded hover:bg-indigo-100 transition-colors"
            >
              該当を全選択
            </button>
            <button
              onClick={clearSelectedCodes}
              className="text-[10px] font-bold text-slate-600 bg-slate-100 px-2 py-0.5 rounded hover:bg-slate-200 transition-colors"
            >
              選択解除
            </button>
            <span className="text-[10px] text-slate-500 ml-auto">
              選択中 <span className="font-bold text-slate-700">{selectedCodes.length}</span> 件
            </span>
          </div>
        </div>

        {!salesData ? (
          <div className="p-4 text-xs text-slate-500">販売実績データが未取り込みです。</div>
        ) : allEntries.length === 0 ? (
          <div className="p-4 text-xs text-slate-500">該当する販売実績データが見つかりません。</div>
        ) : !normalizedQuery && scopedEntries.length === 0 && isScopedToPage ? (
          <div className="p-4 text-xs text-slate-500">このページ内に介援隊コードがありません。</div>
        ) : filteredEntries.length === 0 ? (
          <div className="p-4 text-xs text-slate-500">該当する介援隊コードが見つかりませんでした。</div>
        ) : (
          <div className="flex-1 min-h-0 flex flex-col">
            <div className="h-44 overflow-y-auto border-b border-slate-100">
              {filteredEntries.map((entry) => (
                <button
                  key={entry.code}
                  onClick={() => toggleSelectedCode(entry.code)}
                  className={`w-full text-left px-2.5 py-2 border-b border-slate-100 transition-colors ${selectedCodeSet.has(entry.code) ? 'bg-indigo-50' : 'hover:bg-slate-50'}`}
                >
                  <div className="flex items-start gap-2">
                    {selectedCodeSet.has(entry.code) ? (
                      <CheckSquare size={14} className="text-indigo-600 mt-0.5 flex-shrink-0" />
                    ) : (
                      <Square size={14} className="text-slate-400 mt-0.5 flex-shrink-0" />
                    )}
                    <div className="min-w-0">
                      <p className="text-[16px] font-mono font-bold text-slate-700 leading-tight">{entry.code}</p>
                      <p className="text-[13px] text-slate-600 mt-0.5 font-medium leading-tight">
                        商品 {entry.items.length} / 数量 <span className="text-[15px] font-black text-indigo-600">{entry.total.toLocaleString()}</span>
                      </p>
                    </div>
                  </div>
                </button>
              ))}
            </div>

            <div className="flex-1 min-h-0 overflow-y-auto p-2.5">
              {selectedEntries.length === 0 ? (
                <div className="text-xs text-slate-500 p-2">表示したい介援隊コードを選択してください。</div>
              ) : (
                <div className="space-y-3">
                  {selectedEntries.map((entry) => (
                    <section key={`selected-${entry.code}`} className="rounded-lg border border-slate-200 bg-white overflow-hidden">
                      <div className="flex items-end justify-between gap-3 border-b border-slate-100 px-2 py-1.5 bg-slate-50">
                        <div>
                          <p className="text-[9px] text-slate-400 uppercase tracking-wider font-bold">Code</p>
                          <p className="text-[22px] font-mono font-black text-slate-800 leading-none">{entry.code}</p>
                        </div>
                        <div className="text-right">
                          <p className="text-[9px] text-slate-400 uppercase font-bold">Total</p>
                          <p className="text-[24px] font-black text-indigo-600 font-mono leading-none">{entry.total.toLocaleString()}</p>
                        </div>
                      </div>
                      <ul className="space-y-1.5 p-2">
                        {entry.items.map((item, index) => (
                          <li key={`${entry.code}-${index}`} className="flex justify-between items-center gap-2 p-2 rounded-md border border-slate-100 hover:bg-slate-50 transition-colors">
                            <div className="min-w-0">
                              <p className="text-[14px] font-bold text-slate-700 truncate leading-tight">{item.name || '-'}</p>
                              <p className="text-[12px] text-slate-500 truncate">{item.spec || '-'}</p>
                            </div>
                            <span className="text-[15px] font-mono font-black text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded">
                              {parseInt(item.count) || 0}
                            </span>
                          </li>
                        ))}
                      </ul>
                    </section>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
});

export default SalesCodeLookupModal;
