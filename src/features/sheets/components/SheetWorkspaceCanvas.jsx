import { Check, ChevronLeft, ChevronRight } from 'lucide-react';

import Sheet from './Sheet';

const SheetWorkspaceCanvas = ({
  viewMode,
  isTwoPageMode,
  zoomScale,
  displaySheets,
  sheets,
  pageSelection,
  navigation,
  arrange,
  editing,
  sales,
  imageDataById
}) => (
  <div
    data-two-page-workspace={isTwoPageMode ? 'true' : undefined}
    className={`relative z-10 ${viewMode === 'overview'
      ? 'grid grid-cols-2 gap-8 md:grid-cols-3'
      : isTwoPageMode
        ? 'flex min-w-max flex-row items-start justify-center gap-6 pb-32'
        : 'flex flex-col items-center gap-12 pb-32'}`}
    style={{
      transform: `scale(${zoomScale})`,
      transformOrigin: 'top center',
      minHeight: zoomScale > 1 ? `${zoomScale * 100}%` : 'auto'
    }}
  >
    {displaySheets.map((sheet) => {
      const pageIndex = sheets.findIndex((candidate) => candidate.id === sheet.id);
      const isPageSelected = pageSelection.selectedIds.has(sheet.id);
      const panelArrangeView = arrange.workspaceView?.viewsBySheetId?.[sheet.id] || null;
      const isArrangeSheet = arrange.sheetIds.includes(sheet.id) && !!panelArrangeView;
      const renderedPanels = isArrangeSheet ? panelArrangeView.panels : sheet.panels;

      return (
        <div
          key={sheet.id}
          data-two-page-role={isTwoPageMode
            ? (sheet.id === navigation.activeSheetId ? 'primary' : 'secondary')
            : undefined}
          className={`relative group transition-transform duration-300 ${pageSelection.isEnabled ? 'cursor-pointer' : ''} ${isPageSelected ? 'scale-[1.02]' : ''}`}
          onClick={() => {
            if (pageSelection.isEnabled) {
              pageSelection.onToggle(sheet.id);
            } else if (viewMode === 'overview') {
              navigation.onOpenSheet(sheet.id);
            }
          }}
        >
          {pageSelection.isEnabled && (
            <div className={`absolute -inset-4 z-50 rounded-2xl border-4 transition-all duration-200 pointer-events-none ${isPageSelected ? 'border-indigo-500 bg-indigo-500/5 shadow-2xl' : 'border-transparent hover:border-slate-300'}`}>
              <div className={`absolute right-0 top-0 flex h-8 w-8 translate-x-2 -translate-y-2 transform items-center justify-center rounded-full border-2 bg-white shadow-md transition-all ${isPageSelected ? 'scale-110 border-indigo-500 bg-indigo-500 text-white' : 'border-slate-300 text-slate-300'}`}>
                {isPageSelected && <Check size={18} strokeWidth={3} />}
              </div>
            </div>
          )}

          <div className="flex flex-col gap-0">
            <div className={`relative ${pageSelection.isEnabled ? 'pointer-events-none' : ''}`}>
              {viewMode === 'single' && sheet.id === navigation.activeSheetId && (
                <>
                  <button
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation();
                      navigation.onNavigate('prev');
                    }}
                    disabled={navigation.currentIndex <= 0}
                    className={`absolute left-2 top-1/2 z-50 flex h-11 w-11 -translate-y-1/2 items-center justify-center rounded-full border bg-white/95 shadow-lg backdrop-blur transition-all sm:-left-14 ${navigation.currentIndex > 0
                      ? 'border-slate-200 text-slate-600 hover:scale-105 hover:bg-indigo-50 hover:text-indigo-700'
                      : 'cursor-not-allowed border-slate-100 text-slate-300 opacity-40'}`}
                    title="前のページ"
                    aria-label="前のページ"
                  >
                    <ChevronLeft size={24} />
                  </button>

                  <button
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation();
                      navigation.onNavigate('next');
                    }}
                    disabled={navigation.currentIndex === -1 || navigation.currentIndex >= navigation.totalCount - 1}
                    className={`absolute right-2 top-1/2 z-50 flex h-11 w-11 -translate-y-1/2 items-center justify-center rounded-full border bg-white/95 shadow-lg backdrop-blur transition-all sm:-right-14 ${navigation.currentIndex !== -1 && navigation.currentIndex < navigation.totalCount - 1
                      ? 'border-slate-200 text-slate-600 hover:scale-105 hover:bg-indigo-50 hover:text-indigo-700'
                      : 'cursor-not-allowed border-slate-100 text-slate-300 opacity-40'}`}
                    title="次のページ"
                    aria-label="次のページ"
                  >
                    <ChevronRight size={24} />
                  </button>
                </>
              )}

              <Sheet
                sheet={sheet}
                index={pageIndex}
                pageNumber={pageIndex + 1}
                panels={renderedPanels}
                updatePanel={editing.updatePanel}
                isOverview={viewMode === 'overview'}
                zoomScale={zoomScale}
                selection={editing.selection}
                onSelectPanel={editing.isMergeMode ? editing.onSelectPanel : undefined}
                onDeleteSheet={editing.onDeleteSheet}
                highlightEmpty={editing.highlightEmpty}
                highlightLabels={editing.highlightLabels}
                onApplyDragPayloadToPanel={editing.onApplyDragPayloadToPanel}
                onStartPointerDrag={editing.onStartPointerDrag}
                isSalesMode={sales.isMode}
                salesData={sales.data}
                onHoverSales={sales.onHover}
                onLeaveSales={sales.onLeave}
                imageDataById={imageDataById}
                isLabelMode={editing.isLabelMode}
                onChangeGenre={(genreId) => editing.onChangeGenre(sheet.id, genreId)}
                onPreviewImage={editing.onPreviewImage}
                isArrangeMode={isArrangeSheet}
                arrangeDraggingTokenId={arrange.draggingTokenId}
                arrangeAssignedTokenIdsByPanel={isArrangeSheet ? panelArrangeView.assignedTokenIdsByPanel : {}}
                arrangePlacedPanelIndices={isArrangeSheet ? panelArrangeView.placedPanelIndices : new Set()}
                arrangeFloatingTokensByPanel={isArrangeSheet ? panelArrangeView.floatingTokensByPanel : {}}
                onStartArrangeHold={arrange.onStartHold}
                onCancelArrangeHold={arrange.onCancelHold}
                onArrangeDragStateChange={arrange.onDragStateChange}
              />
            </div>
          </div>
        </div>
      );
    })}
  </div>
);

export default SheetWorkspaceCanvas;
