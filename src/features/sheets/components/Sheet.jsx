import React from 'react';

import { GENRES } from '../../../constants/layout';
import Panel from './Panel';

const Sheet = React.memo(({
  sheet,
  pageNumber,
  panels,
  updatePanel,
  isOverview,
  isExportMode = false,
  zoomScale,
  selection,
  onSelectPanel,
  highlightEmpty,
  highlightLabels,
  onApplyDragPayloadToPanel,
  onStartPointerDrag,
  isSalesMode,
  salesData,
  onHoverSales,
  onLeaveSales,
  imageDataById,
  isLabelMode,
  onChangeGenre,
  onPreviewImage,
  isArrangeMode = false,
  arrangeDraggingPanelKey = null,
  onStartArrangeHold,
  onCancelArrangeHold,
  onArrangeDragStateChange
}) => {
  const genre = GENRES.find((candidate) => candidate.id === sheet.genre) || GENRES[0];

  let visibleCounter = 0;
  const displayNumbers = {};
  for (let panelIndex = 0; panelIndex < 16; panelIndex++) {
    const panel = panels[panelIndex];
    if (!panel?.hidden) {
      const isNonCounted = panel.label === '埋草' || panel.label === 'タイトル';
      if (!isNonCounted) {
        visibleCounter++;
        displayNumbers[panelIndex] = visibleCounter;
      } else {
        displayNumbers[panelIndex] = null;
      }
    }
  }

  return (
    <div
      data-pdf-export-sheet-id={isExportMode ? sheet.id : undefined}
      className={`border transition-all duration-300 overflow-hidden ${isExportMode ? '' : (isOverview ? 'hover:shadow-lg hover:scale-[1.02] cursor-pointer' : 'shadow-xl')}`}
      style={{
        width: isExportMode ? '210mm' : (isOverview ? '100%' : `${210 * zoomScale}mm`),
        height: isExportMode ? '297mm' : (isOverview ? 'auto' : `${297 * zoomScale}mm`),
        aspectRatio: '210/297',
        margin: '0 auto',
        display: 'flex',
        flexDirection: 'column',
        borderRadius: 'var(--m3-shape-corner-lg)',
        background: 'var(--m3-surface)',
        borderColor: 'var(--m3-outline-variant)'
      }}
    >
      <div
        className={`${isOverview ? 'h-6 px-2' : 'h-3'} w-full flex-shrink-0 flex items-center justify-between`}
        style={{ backgroundColor: genre.color }}
        title={genre.label}
      >
        {isOverview && (
          <>
            <span className="text-[11px] font-extrabold leading-none text-slate-700/90">
              P.{pageNumber}
            </span>
            <select
              value={sheet.genre}
              onChange={(event) => onChangeGenre?.(event.target.value)}
              onClick={(event) => event.stopPropagation()}
              className="min-w-0 max-w-[70%] cursor-pointer border-0 bg-transparent p-0 text-right text-[11px] font-bold leading-none text-slate-700 shadow-none outline-none focus:ring-0"
              aria-label={`P.${pageNumber}のジャンル`}
            >
              {GENRES.map((candidate) => (
                <option key={candidate.id} value={candidate.id}>{candidate.label}</option>
              ))}
            </select>
          </>
        )}
      </div>

      <div className="flex-1 grid grid-cols-4 grid-rows-4 w-full h-full border-b border-r" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}>
        {Array.from({ length: 16 }).map((_, panelIndex) => {
          const panelData = panels[panelIndex] || {};
          if (panelData.hidden) return null;

          const panelNumber = displayNumbers[panelIndex];
          const isSelected = !isOverview
            && selection?.sheetId === sheet.id
            && selection.indices.includes(panelIndex);

          return (
            <Panel
              key={panelIndex}
              index={panelIndex}
              data={panelData}
              globalNumber={panelNumber}
              onUpdate={(newData) => updatePanel(sheet.id, panelIndex, newData)}
              onUpdateByIndex={(targetIndex, newData) => updatePanel(sheet.id, targetIndex, newData)}
              panels={panels}
              isOverview={isOverview}
              isExportMode={isExportMode}
              isSelected={isSelected}
              onSelect={() => onSelectPanel && onSelectPanel(sheet.id, panelIndex)}
              highlightEmpty={highlightEmpty}
              highlightLabels={highlightLabels}
              sheetId={sheet.id}
              onApplyDragPayloadToPanel={onApplyDragPayloadToPanel}
              onStartPointerDrag={onStartPointerDrag}
              isSalesMode={isSalesMode}
              salesData={salesData}
              onHoverSales={onHoverSales}
              onLeaveSales={onLeaveSales}
              imageDataById={imageDataById}
              isLabelMode={isLabelMode}
              onPreviewImage={onPreviewImage}
              isArrangeMode={isArrangeMode}
              isArrangeDragging={arrangeDraggingPanelKey === `${sheet.id}:${panelIndex}`}
              onStartArrangeHold={onStartArrangeHold}
              onCancelArrangeHold={onCancelArrangeHold}
              onArrangeDragStateChange={onArrangeDragStateChange}
            />
          );
        })}
      </div>

    </div>
  );
});

export default Sheet;
