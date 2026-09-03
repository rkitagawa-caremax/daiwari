import { FileImage, FileSpreadsheet, Loader2 } from 'lucide-react';

import { countPdfCropManualRects, PDF_CROP_RESIZE_HANDLES } from '../../domain/pdfCropEditor';
import { normalizePdfCropCode, pdfTextItemsContainCodeInRect } from '../../domain/pdfCropImport';

const HANDLE_POSITION = Object.freeze({
  nw: '-left-1.5 -top-1.5 cursor-nwse-resize',
  n: 'left-1/2 -top-1.5 -translate-x-1/2 cursor-ns-resize',
  ne: '-right-1.5 -top-1.5 cursor-nesw-resize',
  w: '-left-1.5 top-1/2 -translate-y-1/2 cursor-ew-resize',
  e: '-right-1.5 top-1/2 -translate-y-1/2 cursor-ew-resize',
  sw: '-left-1.5 -bottom-1.5 cursor-nesw-resize',
  s: 'left-1/2 -bottom-1.5 -translate-x-1/2 cursor-ns-resize',
  se: '-right-1.5 -bottom-1.5 cursor-nwse-resize'
});

export const PdfCropFrameOverlay = ({ entry, textItems, isSelected, onPointerDown, onKeyDown }) => {
  const { row, rect, isManual } = entry;
  const code = normalizePdfCropCode(row.code);
  const hasCode = pdfTextItemsContainCodeInRect(textItems, code, rect);
  const tone = isManual
    ? 'border-indigo-500 bg-indigo-400/10'
    : hasCode ? 'border-emerald-500 bg-emerald-400/10' : 'border-amber-500 bg-amber-400/10';
  const badgeTone = hasCode ? 'bg-emerald-600' : 'bg-amber-500';

  return (
    <div
      role="button"
      tabIndex={0}
      aria-label={`${code} の切り抜き枠`}
      title={`${code}｜ドラッグで移動・つまみでサイズ変更・Ctrl+クリックで複数選択`}
      onPointerDown={(event) => onPointerDown(event, entry, 'move')}
      onKeyDown={(event) => onKeyDown(event, entry)}
      className={`absolute touch-none select-none border-2 focus:outline-none ${tone} ${isSelected ? 'z-20 cursor-move ring-2 ring-indigo-400' : 'z-10 cursor-pointer'}`}
      style={{ left: `${rect.x * 100}%`, top: `${rect.y * 100}%`, width: `${rect.width * 100}%`, height: `${rect.height * 100}%` }}
    >
      <span className={`pointer-events-none absolute left-0.5 top-0.5 rounded-md px-1.5 py-0.5 text-[10px] font-black text-white shadow ${badgeTone}`}>{code}{isManual ? '・手動' : ''}</span>
      {isSelected && PDF_CROP_RESIZE_HANDLES.map((handle) => (
        <span
          key={handle}
          role="presentation"
          onPointerDown={(event) => onPointerDown(event, entry, handle)}
          className={`absolute h-3 w-3 rounded-full border-2 border-white bg-indigo-600 shadow ${HANDLE_POSITION[handle]}`}
        />
      ))}
    </div>
  );
};

const PdfCropSourceButton = ({ isSelected, selectedTone, onClick, disabled, icon, title, detail }) => (
  <button
    type="button"
    onClick={onClick}
    disabled={disabled}
    className={`flex w-full min-w-0 items-center gap-2.5 rounded-2xl border px-3 py-2.5 text-left transition-colors disabled:opacity-50 ${isSelected ? selectedTone : 'border-dashed border-slate-300 bg-white hover:bg-slate-50'}`}
  >
    {icon}
    <span className="min-w-0">
      <span className="block truncate text-xs font-black text-slate-700">{title}</span>
      <span className="block truncate text-[10px] text-slate-500">{detail}</span>
    </span>
  </button>
);

export const PdfCropSourceControls = ({
  hasPdf,
  isReadingPdfs,
  isBusy,
  pdfFileCount,
  pdfPageCount,
  maxPages,
  csvFile,
  pdfInputRef,
  csvInputRef,
  onPdfChange,
  onCsvChange
}) => (
  <div className="space-y-2">
    <PdfCropSourceButton
      isSelected={hasPdf}
      selectedTone="border-indigo-300 bg-indigo-50"
      onClick={() => pdfInputRef.current?.click()}
      disabled={isBusy}
      icon={isReadingPdfs
        ? <Loader2 size={18} className="shrink-0 animate-spin text-indigo-600" />
        : <FileImage size={18} className="shrink-0 text-indigo-600" />}
      title={hasPdf ? `PDF ${pdfFileCount}ファイル・${pdfPageCount}ページ` : '校正PDFを選ぶ'}
      detail={isReadingPdfs ? 'PDFを読んでいます…' : hasPdf ? '選び直す' : `まとめて選べます（最大${maxPages}ページ）`}
    />
    <PdfCropSourceButton
      isSelected={!!csvFile}
      selectedTone="border-emerald-300 bg-emerald-50"
      onClick={() => csvInputRef.current?.click()}
      disabled={isBusy}
      icon={<FileSpreadsheet size={18} className="shrink-0 text-emerald-600" />}
      title={csvFile ? '台割CSV' : '台割CSVを選ぶ'}
      detail={csvFile?.name || '台割から書き出した全データCSV'}
    />
    <input ref={pdfInputRef} type="file" accept="application/pdf,.pdf" multiple className="hidden" onChange={onPdfChange} />
    <input ref={csvInputRef} type="file" accept=".csv,text/csv" className="hidden" onChange={onCsvChange} />
  </div>
);

export const PdfCropImportSummary = ({
  summary,
  skipExistingCodes,
  onSkipExistingCodesChange,
  isImporting,
  errorMessage,
  warnings
}) => (
  <section className="rounded-xl border border-slate-200 bg-white px-2.5 py-1.5">
    <p className="flex items-baseline justify-between gap-2">
      <span className="text-[10px] font-bold text-slate-500">保存するコマ</span>
      <span className="text-[10px] font-bold text-slate-500"><span className="text-base font-black text-indigo-700">{summary.importCount}</span>コマ（{summary.pageCount}ページ）</span>
    </p>
    {summary.existingCount > 0 && (
      <label className="mt-1 flex items-start gap-1.5 text-[10px] font-bold leading-snug text-slate-600">
        <input type="checkbox" checked={skipExistingCodes} onChange={onSkipExistingCodesChange} disabled={isImporting} className="mt-0.5" />
        <span>ライブラリにある{summary.existingCount}件は保存しない</span>
      </label>
    )}
    {(errorMessage || warnings.length > 0) && (
      <div className="mt-1 space-y-0.5 text-[10px] leading-snug text-amber-800">
        {errorMessage && <p className="font-bold text-rose-700">{errorMessage}</p>}
        {warnings.map((warning) => <p key={warning}>・{warning}</p>)}
      </div>
    )}
  </section>
);

export const PdfCropManualControls = ({
  hasPdf,
  totalManualCount,
  pageManualCount,
  isImporting,
  onResetPageFrames
}) => {
  if (!hasPdf) return null;
  return (
    <div className="space-y-1">
      <details className="rounded-xl border border-slate-200 bg-white px-2.5 py-1.5">
        <summary className="cursor-pointer text-[10px] font-bold text-slate-500">コマ枠の手動調整{totalManualCount > 0 ? `・${totalManualCount}コマ` : ''}</summary>
        <p className="mt-1 text-[10px] leading-snug text-slate-500">枠をクリック → ドラッグで移動、つまみでサイズ変更、矢印キーで微調整（Shiftで大きく）。Ctrl＋クリックで複数選択して、まとめて動かせます</p>
      </details>
      {pageManualCount > 0 && (
        <button type="button" onClick={onResetPageFrames} disabled={isImporting} className="w-full rounded-lg border border-slate-300 bg-white px-2 py-1 text-[10px] font-bold text-slate-600 hover:bg-slate-50 disabled:opacity-40">このページの{pageManualCount}件を自動に戻す</button>
      )}
    </div>
  );
};

export const PdfCropPageList = ({
  pagePlans,
  csvFile,
  activePageId,
  availablePages,
  pdfSources,
  manualRects,
  isImporting,
  onSelectPage,
  onCatalogPageChange
}) => (
  <section className="min-h-0 flex-1 overflow-auto rounded-2xl border border-slate-200 bg-white p-2">
    <p className="px-2 pb-1 pt-1 text-[10px] font-bold text-slate-500" title="クリックでそのページを表示します。P.○○ を選ぶと、そのPDFに当てるカタログのページを変えられます。">ページ一覧（クリックで表示）</p>
    {pagePlans.length === 0 && <p className="px-2 py-3 text-xs text-slate-500">PDFを選ぶとここに並びます</p>}
    <ul className="space-y-1">
      {pagePlans.map((plan) => {
        const isActive = plan.page.id === activePageId;
        const pageOptions = availablePages.includes(plan.page.catalogPage) ? availablePages : [plan.page.catalogPage, ...availablePages];
        const countLabel = !csvFile ? '' : !plan.hasCsvRows ? 'CSVに該当なし' : `${plan.importRows.length}コマ`;
        const countClass = !csvFile ? 'text-slate-400' : !plan.hasCsvRows ? 'text-rose-600' : plan.importRows.length === 0 ? 'text-amber-600' : 'text-emerald-600';
        const manualCount = countPdfCropManualRects(manualRects, plan.page.id);
        return (
          <li key={plan.page.id}>
            <div
              role="button"
              tabIndex={0}
              onClick={() => onSelectPage(plan.page.id)}
              onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onSelectPage(plan.page.id); } }}
              className={`flex items-center gap-2 rounded-xl border px-2.5 py-2 transition-colors ${isActive ? 'border-indigo-300 bg-indigo-50' : 'border-slate-200 hover:bg-slate-50'}`}
            >
              <span className="min-w-0 flex-1 truncate text-[11px] font-bold text-slate-700" title={plan.page.filename}>{plan.page.filename}{pdfSources[plan.page.sourceIndex]?.numPages > 1 ? ` (${plan.page.pdfPageNumber})` : ''}</span>
              {manualCount > 0 && <span className="shrink-0 rounded-full bg-indigo-100 px-1.5 text-[9px] font-black text-indigo-700" title={`手動調整 ${manualCount}コマ`}>手動{manualCount}</span>}
              <select
                value={plan.page.catalogPage}
                onClick={(event) => event.stopPropagation()}
                onChange={(event) => onCatalogPageChange(plan.page.id, event.target.value)}
                disabled={isImporting}
                className="shrink-0 rounded-md border border-slate-300 bg-white px-1 py-0.5 font-mono text-[10px] font-bold text-slate-700"
                aria-label={`${plan.page.filename} の対象ページ`}
              >
                {pageOptions.map((page) => <option key={page} value={page}>P.{page}</option>)}
              </select>
              <span className={`w-14 shrink-0 text-right text-[10px] font-bold ${countClass}`}>{countLabel}</span>
            </div>
          </li>
        );
      })}
    </ul>
  </section>
);
