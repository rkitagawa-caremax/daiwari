import React, { useEffect, useMemo, useRef, useState } from 'react';
import { GripVertical, X } from 'lucide-react';

import {
  FREE_LABEL_COLORS,
  FREE_LABEL_HALF_HEIGHT_PX,
  FREE_LABEL_HALF_WIDTH_PX
} from '../../../constants/layout';
import { normalizeCode } from '../../../domain/productCodes';
import {
  DAIWARI_PANEL_DROPZONE_PREFIX,
  clearActiveNativeDragPayload,
  getDragPayload,
  isDropEventHandled,
  markDropEventHandled,
  setDragPayload
} from '../../../lib/dragPayload';
import { clamp } from '../../../lib/math';

const Panel = React.memo(({
  index,
  data,
  globalNumber,
  onUpdate,
  onUpdateByIndex,
  panels,
  isOverview,
  isSelected,
  onSelect,
  highlightEmpty,
  highlightLabels,
  sheetId,
  onApplyDragPayloadToPanel,
  onStartPointerDrag,
  isSalesMode,
  salesData,
  onHoverSales,
  onLeaveSales,
  imageDataById,
  isLabelMode
}) => {
  const [isHovered, setIsHovered] = useState(false);
  const textareaRef = useRef(null);
  const [localText, setLocalText] = useState(data.text || '');
  const [labelDrafts, setLabelDrafts] = useState({});
  const isFocusedRef = useRef(false);
  const editingLabelIdRef = useRef(null);
  const panelRef = useRef(null);

  const matchedSales = useMemo(() => {
    if (!isSalesMode || !data.code || !salesData) return null;
    const normalizedTarget = normalizeCode(data.code);
    return salesData[normalizedTarget] || null;
  }, [isSalesMode, data.code, salesData]);

  useEffect(() => {
    if (!isFocusedRef.current) {
      setLocalText(data.text || '');
    }
  }, [data.text]);

  useEffect(() => {
    const labels = data.freeLabels || (data.freeText
      ? [{ id: 'legacy', text: data.freeText, x: 50, y: 50, colorIndex: 0 }]
      : []);
    setLabelDrafts((previous) => {
      const next = {};
      labels.forEach((label) => {
        if (editingLabelIdRef.current === label.id && previous[label.id] !== undefined) {
          next[label.id] = previous[label.id];
        } else {
          next[label.id] = label.text || '';
        }
      });
      return next;
    });
  }, [data.freeLabels, data.freeText]);

  const handleMouseEnter = (event) => {
    setIsHovered(true);
    if (isSalesMode && matchedSales && onHoverSales) {
      const rect = event.currentTarget.getBoundingClientRect();
      onHoverSales(matchedSales, { x: rect.right, y: rect.top });
    }
  };

  const handleMouseLeave = () => {
    setIsHovered(false);
    if (isSalesMode && onLeaveSales) {
      onLeaveSales();
    }
  };

  const handleDrop = (event) => {
    event.preventDefault();
    if (isOverview) return;
    const nativeDropEvent = event.nativeEvent;
    if (isDropEventHandled(nativeDropEvent)) return;
    const handled = onApplyDragPayloadToPanel?.(
      sheetId,
      index,
      getDragPayload(event.dataTransfer) || {}
    );
    if (handled) {
      markDropEventHandled(nativeDropEvent);
      clearActiveNativeDragPayload();
    }
  };

  const handleDragStart = (event) => {
    if (!resolvedImage && !data.label && !data.isText && !data.code) {
      event.preventDefault();
      return;
    }

    const currentText = textareaRef.current ? textareaRef.current.value : localText;
    setDragPayload(event.dataTransfer, {
      moveSourceType: 'panel',
      sourceSheetId: sheetId,
      sourceIndex: index,
      textData: currentText || ''
    });
    event.dataTransfer.effectAllowed = 'move';
  };

  const handlePointerDragStart = (event) => {
    if (!onStartPointerDrag) return;
    const currentText = textareaRef.current ? textareaRef.current.value : localText;
    onStartPointerDrag(event, {
      payload: {
        moveSourceType: 'panel',
        sourceSheetId: sheetId,
        sourceIndex: index,
        textData: currentText || ''
      },
      preview: {
        image: resolvedImage || null,
        label: data.label || null,
        code: data.code || null,
        text: data.isText ? (currentText || data.text || '') : ''
      }
    });
  };

  const handleDragOver = (event) => {
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'move';
    }
  };

  const resolvedImage = data.image || (data.imageId ? imageDataById?.[data.imageId] : null);
  const isEmpty = !resolvedImage && !data.label && !data.code && !data.isText;
  const freeLabelsCount = data.freeLabels?.length || 0;
  const hasFreeLabel = freeLabelsCount > 0 || (!!data.freeText && freeLabelsCount === 0);
  const shouldHighlightLabel = isOverview && highlightLabels && hasFreeLabel;
  const shouldHighlightEmpty = highlightEmpty && (!resolvedImage && (isEmpty || !!data.code));
  const mergeSelectionGlow = isSelected
    ? '0 0 0 2px rgba(37, 99, 235, 0.88), inset 0 0 0 1px rgba(191, 219, 254, 0.95), 0 0 24px rgba(59, 130, 246, 0.55)'
    : null;

  const handleFocusTextarea = () => {
    if (textareaRef.current) {
      textareaRef.current.focus();
    }
  };

  const handleTextChange = (event) => {
    setLocalText(event.target.value);
  };

  const handleTextBlur = () => {
    isFocusedRef.current = false;
    if (localText !== data.text) {
      onUpdate({ ...data, text: localText });
    }
  };

  const getLabelStyle = (label) => {
    switch (label) {
      case '新規商品':
      case '新規商品未確定': return { bg: 'var(--dummy-gray)', text: 'var(--dummy-text-color)' };
      case 'テキスト': return { bg: 'var(--dummy-volt)', text: 'var(--dummy-text-color)' };
      case 'タイトル': return { bg: 'var(--dummy-red)', text: 'var(--dummy-text-color)' };
      case '埋草': return { bg: 'var(--dummy-green)', text: 'var(--dummy-text-color)' };
      default: return { bg: 'var(--m3-surface-container)', text: 'var(--m3-on-surface)' };
    }
  };

  const handlePanelClick = (event) => {
    if (isOverview) return;

    if (isLabelMode) {
      event.stopPropagation();
      if (!panelRef.current) return;

      const currentRect = panelRef.current.getBoundingClientRect();
      let targetIndex = index;
      let targetPanelElement = panelRef.current;

      const overflowDirections = [];
      if ((event.clientX - currentRect.left) < FREE_LABEL_HALF_WIDTH_PX) overflowDirections.push('left');
      if ((currentRect.right - event.clientX) < FREE_LABEL_HALF_WIDTH_PX) overflowDirections.push('right');
      if ((event.clientY - currentRect.top) < FREE_LABEL_HALF_HEIGHT_PX) overflowDirections.push('up');
      if ((currentRect.bottom - event.clientY) < FREE_LABEL_HALF_HEIGHT_PX) overflowDirections.push('down');

      if (overflowDirections.length > 0 && typeof document !== 'undefined') {
        const idPrefix = `panel-${sheetId}-`;
        const viewportMaxX = Math.max(0, window.innerWidth - 1);
        const viewportMaxY = Math.max(0, window.innerHeight - 1);

        for (const direction of overflowDirections) {
          let probeX = event.clientX;
          let probeY = event.clientY;

          if (direction === 'left') probeX = currentRect.left - 1;
          if (direction === 'right') probeX = currentRect.right + 1;
          if (direction === 'up') probeY = currentRect.top - 1;
          if (direction === 'down') probeY = currentRect.bottom + 1;

          const candidate = document.elementFromPoint(
            clamp(probeX, 0, viewportMaxX),
            clamp(probeY, 0, viewportMaxY)
          );
          const panelElement = candidate?.closest?.(`[id^="${idPrefix}"]`);
          if (!panelElement || panelElement === panelRef.current) continue;

          const rawIndex = (panelElement.id || '').slice(idPrefix.length);
          const parsedIndex = Number.parseInt(rawIndex, 10);
          if (Number.isNaN(parsedIndex)) continue;
          if (panels?.[parsedIndex]?.hidden) continue;

          targetIndex = parsedIndex;
          targetPanelElement = panelElement;
          break;
        }
      }

      const targetRect = targetPanelElement.getBoundingClientRect();
      const rawXPercent = ((event.clientX - targetRect.left) / targetRect.width) * 100;
      const rawYPercent = ((event.clientY - targetRect.top) / targetRect.height) * 100;

      const minXPercent = (FREE_LABEL_HALF_WIDTH_PX / targetRect.width) * 100;
      const maxXPercent = 100 - minXPercent;
      const minYPercent = (FREE_LABEL_HALF_HEIGHT_PX / targetRect.height) * 100;
      const maxYPercent = 100 - minYPercent;

      const xPercent = minXPercent <= maxXPercent
        ? clamp(rawXPercent, minXPercent, maxXPercent)
        : 50;
      const yPercent = minYPercent <= maxYPercent
        ? clamp(rawYPercent, minYPercent, maxYPercent)
        : 50;

      const targetPanelData = targetIndex === index ? data : (panels?.[targetIndex] || {});
      const previousLabels = targetPanelData.freeLabels || [];
      const migratedLabels = targetPanelData.freeText && previousLabels.length === 0
        ? [{ id: 'legacy', text: targetPanelData.freeText, x: 50, y: 50, colorIndex: 0 }]
        : previousLabels;

      const colorIndex = migratedLabels.length % FREE_LABEL_COLORS.length;
      const newLabel = {
        id: Date.now().toString() + Math.random().toString(36).substring(2, 7),
        text: 'ラベル',
        x: xPercent,
        y: yPercent,
        colorIndex
      };

      const nextPanelData = {
        ...targetPanelData,
        freeText: null,
        freeLabels: [...migratedLabels, newLabel]
      };

      if (targetIndex === index || !onUpdateByIndex) {
        onUpdate(nextPanelData);
      } else {
        onUpdateByIndex(targetIndex, nextPanelData);
      }
      return;
    }

    if (onSelect) onSelect();
  };

  const labelStyle = data.label ? getLabelStyle(data.label) : {};
  const salesTotal = matchedSales
    ? matchedSales.reduce((total, item) => total + (parseInt(item.count) || 0), 0)
    : 0;

  return (
    <div
      ref={panelRef}
      id={`panel-${sheetId}-${index}`}
      data-daiwari-dropzone-id={!isOverview ? `${DAIWARI_PANEL_DROPZONE_PREFIX}${sheetId}:${index}` : undefined}
      className={`relative border-t border-l flex flex-col items-center justify-center overflow-hidden transition-all duration-300
        ${(shouldHighlightLabel || shouldHighlightEmpty) ? 'ring-inset ring-2' : 'hover:shadow-md hover:z-10'}
        ${isSelected ? 'ring-4 z-20 shadow-xl' : ''}
        ${!isEmpty && !isOverview && !data.isText ? 'cursor-grab active:cursor-grabbing' : ''}
        ${isSalesMode ? 'hover:ring-4 hover:z-40' : ''}
      `}
      draggable={!isEmpty && !isOverview && !isLabelMode && !data.isText}
      onDragStart={handleDragStart}
      onDropCapture={handleDrop}
      onDragOverCapture={handleDragOver}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onPointerDown={!isEmpty && !isOverview && !isLabelMode && !data.isText ? handlePointerDragStart : undefined}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      onClick={handlePanelClick}
      style={{
        ...(shouldHighlightEmpty ? {} : {}),
        '--tw-ring-color': isSelected
          ? '#3b82f6'
          : shouldHighlightLabel
            ? '#22c55e'
            : shouldHighlightEmpty
              ? 'var(--m3-error)'
              : 'transparent',
        cursor: isLabelMode
          ? 'crosshair'
          : (isOverview ? 'pointer' : (!isEmpty && !data.isText ? 'cursor-grab' : 'default')),
        gridColumn: `span ${data.colSpan || 1}`,
        gridRow: `span ${data.rowSpan || 1}`,
        borderColor: isSelected
          ? '#3b82f6'
          : shouldHighlightLabel
            ? '#16a34a'
            : 'var(--m3-outline-variant)',
        backgroundColor: shouldHighlightLabel
          ? 'rgba(34, 197, 94, 0.16)'
          : shouldHighlightEmpty
            ? 'var(--m3-error-container)'
            : isSalesMode && matchedSales
              ? 'var(--m3-secondary-container)'
              : 'var(--m3-surface)',
        boxShadow: isSelected
          ? mergeSelectionGlow
          : shouldHighlightLabel
            ? '0 0 0 1.5px rgba(22, 163, 74, 0.65), inset 0 0 0 1px rgba(34, 197, 94, 0.55), 0 0 20px rgba(34, 197, 94, 0.35)'
            : undefined,
        touchAction: !isEmpty && !isOverview && !isLabelMode ? 'none' : undefined,
      }}
    >
      {resolvedImage && (
        <img
          src={resolvedImage}
          alt="content"
          className="w-full h-full object-contain absolute inset-0 z-0 pointer-events-none"
          loading="lazy"
          decoding="async"
          draggable={false}
        />
      )}

      {(() => {
        const labels = data.freeLabels || (data.freeText
          ? [{ id: 'legacy', text: data.freeText, x: 50, y: 50, colorIndex: 0 }]
          : []);
        if (labels.length === 0) return null;

        return labels.map((label, labelIndex) => {
          const color = FREE_LABEL_COLORS[label.colorIndex % FREE_LABEL_COLORS.length];
          const draftText = labelDrafts[label.id] ?? label.text ?? '';
          return (
            <div
              key={label.id}
              className="absolute z-[25] transform -translate-x-1/2 -translate-y-1/2"
              style={{ left: `${label.x}%`, top: `${label.y}%`, minWidth: '80px', maxWidth: '90%' }}
              onClick={(event) => event.stopPropagation()}
              onMouseDown={(event) => event.stopPropagation()}
            >
              <div className="relative group flex items-center justify-center">
                <textarea
                  className="w-full bg-white/95 backdrop-blur-sm rounded-lg p-1.5 text-xs font-bold text-center resize-none focus:outline-none shadow-lg min-h-[2.5em] overflow-hidden transition-shadow focus:ring-2"
                  style={{
                    border: `2px solid ${color.border}`,
                    color: '#334155',
                    '--tw-ring-color': color.border
                  }}
                  value={draftText}
                  onChange={(event) => {
                    const nextText = event.target.value;
                    setLabelDrafts((previous) => ({ ...previous, [label.id]: nextText }));
                  }}
                  onFocus={() => { editingLabelIdRef.current = label.id; }}
                  onBlur={(event) => {
                    const committedText = (event.target.value || '').toString();
                    editingLabelIdRef.current = null;
                    if (committedText === (label.text || '')) return;
                    const newLabels = [...labels];
                    newLabels[labelIndex] = { ...label, text: committedText };
                    onUpdate({ ...data, freeLabels: newLabels, freeText: null });
                  }}
                  placeholder="入力"
                  rows={Math.max(1, draftText.split('\n').length)}
                />
                {!isOverview && (
                  <button
                    onClick={(event) => {
                      event.stopPropagation();
                      const newLabels = labels.filter((_, indexToKeep) => indexToKeep !== labelIndex);
                      onUpdate({ ...data, freeLabels: newLabels, freeText: null });
                    }}
                    className="absolute -top-2 -right-2 text-white rounded-full p-1 shadow-md transition-colors opacity-0 group-hover:opacity-100 group-focus-within:opacity-100 z-10"
                    style={{ backgroundColor: color.border }}
                    title="ラベルを削除"
                  >
                    <X size={12} strokeWidth={3} />
                  </button>
                )}
              </div>
            </div>
          );
        });
      })()}

      {isSalesMode && matchedSales && (
        <div className="absolute inset-0 z-30 bg-black/60 flex flex-col p-2 text-white pointer-events-none">
          <div className="flex justify-between items-start mb-1">
            <span className="text-[10px] bg-emerald-500 text-white px-1 py-0.5 rounded font-bold shadow-sm">
              実績
            </span>
            <span className="text-xl font-bold font-mono tracking-tighter text-emerald-300">
              {salesTotal.toLocaleString()}
            </span>
          </div>
          <div className="flex-1 overflow-hidden space-y-1">
            {matchedSales.slice(0, 3).map((item, itemIndex) => (
              <div key={itemIndex} className="flex justify-between items-baseline text-[9px] border-b border-white/20 pb-0.5">
                <span className="truncate w-2/3 opacity-90">{item.name} {item.spec}</span>
                <span className="font-mono font-bold opacity-100">{item.count}</span>
              </div>
            ))}
            {matchedSales.length > 3 && (
              <div className="text-[8px] text-center opacity-70 italic mt-1">
                他 {matchedSales.length - 3} 件...
              </div>
            )}
          </div>
        </div>
      )}

      {!resolvedImage && (!data.code || (highlightEmpty && (!resolvedImage && (isEmpty || !!data.code)))) && !data.label && !isOverview && (
        <div className={`flex flex-col items-center justify-center transition-opacity duration-300 ${(highlightEmpty && (!resolvedImage && (isEmpty || !!data.code))) ? 'opacity-100' : 'opacity-0 hover:opacity-100'}`}>
          <span className="text-[10px] select-none font-bold" style={{ color: (highlightEmpty && isEmpty) ? 'var(--m3-on-error-container)' : 'var(--m3-outline)' }}>
            {highlightEmpty && isEmpty ? '空き' : 'Drop Here'}
          </span>
        </div>
      )}

      {data.label && !data.isText && (
        <div
          className="absolute inset-0 flex items-center justify-center pointer-events-none opacity-90 z-10"
          style={{ background: labelStyle.bg, color: labelStyle.text }}
        >
          <span className="font-bold text-sm">{data.label}</span>
        </div>
      )}

      {data.isText && (
        isOverview ? (
          <div
            className="absolute inset-0 p-2 z-20 flex items-center justify-center text-center overflow-hidden pointer-events-none"
            style={{ background: labelStyle.bg || 'rgba(255,255,255,0.5)' }}
          >
            <p className="text-[8px] leading-tight break-words whitespace-pre-wrap font-bold font-sans" style={{ color: labelStyle.text || 'var(--m3-on-surface)' }}>{data.text}</p>
          </div>
        ) : (
          <div
            data-text-editor="true"
            className="absolute inset-0 z-20 flex items-center justify-center p-4 cursor-text transition-colors hover:brightness-95 focus-within:brightness-100"
            style={{ background: labelStyle.bg || 'rgba(255,255,255,0.4)' }}
            onClick={(event) => { event.stopPropagation(); handleFocusTextarea(); }}
            onMouseDown={(event) => event.stopPropagation()}
          >
            <button
              className="absolute top-1 left-1 rounded-md p-1.5 cursor-grab active:cursor-grabbing border shadow-sm z-30"
              style={{ background: 'var(--m3-surface)', borderColor: 'var(--m3-outline-variant)', color: 'var(--m3-on-surface-variant)', touchAction: 'none' }}
              title="ドラッグして移動"
              draggable
              onMouseDown={(event) => event.stopPropagation()}
              onClick={(event) => event.stopPropagation()}
              onDragStart={handleDragStart}
              onPointerDown={handlePointerDragStart}
            >
              <GripVertical size={12} />
            </button>
            <textarea
              ref={textareaRef}
              className={`w-full bg-transparent resize-none focus:outline-none text-sm font-bold text-center overflow-hidden font-sans placeholder:text-slate-400/70 ${labelStyle.text || 'text-slate-800'}`}
              value={localText}
              onChange={handleTextChange}
              onFocus={() => { isFocusedRef.current = true; }}
              onBlur={handleTextBlur}
              placeholder="テキストを入力"
              rows={Math.max(2, (localText || '').split('\n').length)}
              style={{ maxHeight: '100%' }}
              draggable={false}
              onMouseDown={(event) => event.stopPropagation()}
              onClick={(event) => event.stopPropagation()}
            />
          </div>
        )
      )}

      {globalNumber && !isSalesMode && (
        <div className="absolute top-0 left-0 z-10 flex shadow-sm pointer-events-none opacity-90">
          <div className="text-[10px] px-1.5 py-0.5 font-bold rounded-br-sm shadow-sm" style={{ background: 'var(--m3-surface-variant)', color: 'var(--m3-on-surface-variant)' }}>
            {globalNumber}
          </div>
          {data.code && (
            <div className="text-[10px] px-1.5 py-0.5 font-bold border-r border-b shadow-sm font-mono" style={{ background: 'var(--m3-surface)', color: 'var(--m3-on-surface)', borderColor: 'var(--m3-outline-variant)' }}>
              {data.code}
            </div>
          )}
        </div>
      )}

      {!isOverview && (
        <>
          {(resolvedImage || data.label || data.isText || data.code) && isHovered && !isSalesMode && (
            <button
              onClick={(event) => {
                event.stopPropagation();
                onUpdate({ ...data, image: null, imageId: null, label: null, code: null, isText: false, text: '' });
              }}
              className="absolute top-1.5 right-1.5 rounded-full p-1.5 shadow-lg border z-30 transition-all duration-300 transform hover:scale-110 active:scale-90"
              style={{ background: 'var(--m3-surface)', color: 'var(--m3-error)', borderColor: 'var(--m3-error)' }}
            >
              <X size={14} strokeWidth={3} />
            </button>
          )}
        </>
      )}
    </div>
  );
});

export default Panel;
