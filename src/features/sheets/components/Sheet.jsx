import React from 'react';

import { GENRES } from '../../../constants/layout';
import Panel from './Panel';

const Sheet = React.memo(({
  sheet,
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
  isLabelMode
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
        className="h-3 w-full flex-shrink-0"
        style={{ backgroundColor: genre.color }}
        title={genre.label}
      />

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
            />
          );
        })}
      </div>

    </div>
  );
});

export default Sheet;
