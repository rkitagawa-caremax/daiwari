import React from 'react';

import Sheet from './Sheet';

const NOOP = () => {};

const PdfExportSurface = React.memo(({ page, imageDataById }) => {
  if (!page?.sheet) return null;

  return (
    <div
      aria-hidden="true"
      className="fixed top-0 pointer-events-none"
      style={{ left: '-10000px', width: '210mm', height: '297mm', zIndex: -1 }}
    >
      <Sheet
        sheet={page.sheet}
        panels={page.sheet.panels || []}
        updatePanel={NOOP}
        isOverview={false}
        isExportMode
        zoomScale={1}
        selection={{ sheetId: null, indices: [] }}
        onSelectPanel={undefined}
        highlightEmpty={false}
        highlightLabels={false}
        onApplyDragPayloadToPanel={undefined}
        onStartPointerDrag={undefined}
        isSalesMode={false}
        salesData={null}
        onHoverSales={undefined}
        onLeaveSales={undefined}
        imageDataById={imageDataById}
        isLabelMode={false}
      />
    </div>
  );
});

export default PdfExportSurface;
