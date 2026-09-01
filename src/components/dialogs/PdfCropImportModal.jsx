import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Crop,
  FileImage,
  FileSpreadsheet,
  Loader2,
  X
} from 'lucide-react';

import {
  DEFAULT_PDF_GRID_BOUNDS,
  PDF_CROP_SIZE_OPTIONS,
  getPdfCropRect,
  inferCatalogStartPage,
  isPdfCropRowInsideGrid,
  normalizePdfCropCode,
  parsePdfCropCsv,
  pdfTextItemsContainCode,
  updatePdfCropRowSize
} from '../../domain/pdfCropImport';
import { readFileAutoEncoding } from '../../lib/csv';
import { cropPdfPageToFile, openPdfFile, renderPdfPage } from '../../lib/pdfCropImport';

const BOUND_LABELS = Object.freeze({
  left: '左余白',
  top: '上余白',
  right: '右余白',
  bottom: '下余白'
});

const getExistingCodeSet = (images = []) => new Set(images
  .map((image) => normalizePdfCropCode(image.code || image.name || ''))
  .filter(Boolean));

const buildSelection = (rows, existingCodes) => {
  const seenCodes = new Set();
  return new Set(rows.filter((row) => {
    if (!isPdfCropRowInsideGrid(row) || existingCodes.has(row.code) || seenCodes.has(row.code)) return false;
    seenCodes.add(row.code);
    return true;
  }).map((row) => row.id));
};

const findGridConflicts = (rows) => {
  const conflicts = new Set();
  const pages = new Map();
  rows.forEach((row) => {
    if (!pages.has(row.pageNumber)) pages.set(row.pageNumber, []);
    pages.get(row.pageNumber).push(row);
  });
  pages.forEach((pageRows) => {
    const occupied = new Map();
    pageRows.forEach((row) => {
      if (!isPdfCropRowInsideGrid({ ...row, layoutStatus: 'ready' })) {
        conflicts.add(row.id);
        return;
      }
      for (let y = row.yPos; y < row.yPos + row.rowSpan; y++) {
        for (let x = row.xPos; x < row.xPos + row.colSpan; x++) {
          const key = `${x}:${y}`;
          if (occupied.has(key)) {
            conflicts.add(row.id);
            conflicts.add(occupied.get(key));
          } else {
            occupied.set(key, row.id);
          }
        }
      }
    });
  });
  return conflicts;
};

const PdfCropImportModal = ({ isOpen, onClose, onImport, existingImages = [], isLocked = false }) => {
  const canvasRef = useRef(null);
  const pdfInputRef = useRef(null);
  const csvInputRef = useRef(null);
  const documentRef = useRef(null);
  const renderSequenceRef = useRef(0);
  const [pdfFile, setPdfFile] = useState(null);
  const [csvFile, setCsvFile] = useState(null);
  const [pdfDocument, setPdfDocument] = useState(null);
  const [rows, setRows] = useState([]);
  const [issues, setIssues] = useState([]);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [catalogStartPage, setCatalogStartPage] = useState(1);
  const [previewPage, setPreviewPage] = useState(1);
  const [preview, setPreview] = useState({ isLoading: false, textItems: [], width: 0, height: 0 });
  const [bounds, setBounds] = useState({ ...DEFAULT_PDF_GRID_BOUNDS });
  const [errorMessage, setErrorMessage] = useState('');
  const [isImporting, setIsImporting] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, message: '' });

  const existingCodes = useMemo(() => getExistingCodeSet(existingImages), [existingImages]);
  const conflictIds = useMemo(() => findGridConflicts(rows), [rows]);
  const duplicateCodes = useMemo(() => {
    const counts = new Map();
    rows.forEach((row) => counts.set(row.code, (counts.get(row.code) || 0) + 1));
    return new Set([...counts].filter(([, count]) => count > 1).map(([code]) => code));
  }, [rows]);

  useEffect(() => {
    documentRef.current = pdfDocument;
  }, [pdfDocument]);

  useEffect(() => () => {
    renderSequenceRef.current++;
    documentRef.current?.destroy?.();
  }, []);

  useEffect(() => {
    if (!isOpen || !pdfDocument || !canvasRef.current) return undefined;
    const sequence = ++renderSequenceRef.current;
    let cancelled = false;
    setPreview((current) => ({ ...current, isLoading: true }));
    setErrorMessage('');

    renderPdfPage(pdfDocument, previewPage, { scale: 1.35, canvas: canvasRef.current })
      .then((result) => {
        if (cancelled || sequence !== renderSequenceRef.current) return;
        setPreview({
          isLoading: false,
          textItems: result.textItems,
          width: result.width,
          height: result.height
        });
      })
      .catch((error) => {
        if (cancelled || sequence !== renderSequenceRef.current) return;
        console.error('PDF preview rendering failed:', error);
        setPreview((current) => ({ ...current, isLoading: false }));
        setErrorMessage(`PDFプレビューを表示できません。${error?.message || ''}`);
      });

    return () => {
      cancelled = true;
    };
  }, [isOpen, pdfDocument, previewPage]);

  if (!isOpen) return null;

  const updateSelectionsForRows = (nextRows) => {
    setSelectedIds(buildSelection(nextRows, existingCodes));
  };

  const handlePdfChange = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setErrorMessage('');
    setPreview((current) => ({ ...current, isLoading: true }));
    try {
      const nextDocument = await openPdfFile(file);
      await documentRef.current?.destroy?.();
      documentRef.current = nextDocument;
      setPdfDocument(nextDocument);
      setPdfFile(file);
      setCatalogStartPage(inferCatalogStartPage(file.name));
      setPreviewPage(1);
    } catch (error) {
      console.error('PDF loading failed:', error);
      setPreview((current) => ({ ...current, isLoading: false }));
      setErrorMessage(`PDFを読み込めません。${error?.message || ''}`);
    }
  };

  const handleCsvChange = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setErrorMessage('');
    try {
      const content = await readFileAutoEncoding(file);
      const parsed = parsePdfCropCsv(content);
      setCsvFile(file);
      setRows(parsed.rows);
      setIssues(parsed.issues);
      updateSelectionsForRows(parsed.rows);
      if (parsed.rows.length === 0) setErrorMessage('切り抜き対象の介援隊コードがCSVにありません。');
    } catch (error) {
      console.error('Crop CSV loading failed:', error);
      setErrorMessage(`CSVを読み込めません。${error?.message || ''}`);
    }
  };

  const updateRow = (rowId, patch) => {
    setRows((current) => current.map((row) => {
      if (row.id !== rowId) return row;
      const next = { ...row, ...patch, layoutStatus: 'ready' };
      if (patch.code !== undefined) {
        next.code = String(patch.code).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9_-]/g, '');
        const normalizedCode = normalizePdfCropCode(next.code);
        next.filename = normalizedCode ? `${normalizedCode}.jpg` : '';
      }
      return next;
    }));
  };

  const normalizeRowCode = (rowId) => {
    setRows((current) => current.map((row) => {
      if (row.id !== rowId) return row;
      const code = normalizePdfCropCode(row.code);
      return { ...row, code, filename: code ? `${code}.jpg` : '' };
    }));
  };

  const updateRowSize = (rowId, sizeType) => {
    setRows((current) => current.map((row) => (
      row.id === rowId
        ? { ...updatePdfCropRowSize(row, sizeType), layoutStatus: 'ready' }
        : row
    )));
  };

  const toggleSelection = (rowId) => {
    setSelectedIds((current) => {
      const next = new Set(current);
      if (next.has(rowId)) next.delete(rowId);
      else next.add(rowId);
      return next;
    });
  };

  const currentCatalogPage = catalogStartPage + previewPage - 1;
  const previewRows = rows.filter((row) => row.pageNumber === currentCatalogPage);
  const selectedRows = rows.filter((row) => selectedIds.has(row.id));
  const getPdfPageNumber = (row) => row.pageNumber - catalogStartPage + 1;
  const canImportRow = (row) => (
    !!normalizePdfCropCode(row.code)
    && !conflictIds.has(row.id)
    && getPdfPageNumber(row) >= 1
    && getPdfPageNumber(row) <= (pdfDocument?.numPages || 0)
  );
  const importableRows = selectedRows.filter(canImportRow);

  const handleImport = async () => {
    if (isLocked || isImporting || !pdfDocument) return;
    if (importableRows.length === 0) {
      setErrorMessage('登録できる切り抜き対象を1件以上選択してください。');
      return;
    }

    setIsImporting(true);
    setErrorMessage('');
    setProgress({ current: 0, total: importableRows.length, message: '切り抜きを準備しています…' });
    try {
      const normalizedRows = importableRows.map((row) => {
        const code = normalizePdfCropCode(row.code);
        return { ...row, code, filename: `${code}.jpg` };
      });
      const pageGroups = new Map();
      normalizedRows.forEach((row) => {
        const pageNumber = getPdfPageNumber(row);
        if (!pageGroups.has(pageNumber)) pageGroups.set(pageNumber, []);
        pageGroups.get(pageNumber).push(row);
      });

      const items = [];
      let completed = 0;
      for (const [pageNumber, pageRows] of pageGroups) {
        setProgress((current) => ({ ...current, message: `PDF ${pageNumber}ページ目を高解像度で描画しています…` }));
        const rendered = await renderPdfPage(pdfDocument, pageNumber, { scale: 3 });
        for (const row of pageRows) {
          setProgress({
            current: completed,
            total: importableRows.length,
            message: `${row.code} を切り抜いています…`
          });
          const cropRect = getPdfCropRect(row, bounds);
          const file = await cropPdfPageToFile({
            canvas: rendered.canvas,
            normalizedRect: cropRect,
            filename: row.filename
          });
          items.push({
            file,
            code: row.code,
            sourcePdfName: pdfFile.name,
            sourcePage: row.pageNumber,
            pdfPageNumber: pageNumber,
            sizeType: row.sizeType,
            cropRect
          });
          completed++;
          setProgress({ current: completed, total: importableRows.length, message: `${completed}/${importableRows.length}件を生成しました` });
        }
        rendered.canvas.width = 1;
        rendered.canvas.height = 1;
      }

      const result = await onImport(items, ({ current, total, message }) => {
        setProgress({ current, total, message });
      });
      if (!result || result.successCount > 0) onClose();
    } catch (error) {
      console.error('PDF crop import failed:', error);
      setErrorMessage(`切り抜き画像の登録に失敗しました。${error?.message || ''}`);
    } finally {
      setIsImporting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/55 p-4 backdrop-blur-sm">
      <div className="flex h-[min(920px,94vh)] w-[min(1420px,96vw)] flex-col overflow-hidden rounded-3xl border border-white/30 bg-slate-100 shadow-2xl" role="dialog" aria-modal="true" aria-label="PDF＋CSV画像取り込み">
        <header className="flex items-center justify-between border-b border-slate-200 bg-white px-6 py-4">
          <div>
            <h2 className="flex items-center gap-2 text-lg font-black text-slate-800"><Crop size={20} className="text-indigo-600" />PDF＋CSV画像取り込み</h2>
            <p className="mt-1 text-xs text-slate-500">CSVのページ・座標・コマサイズを使い、前号PDFから商品コマを切り抜いて介援隊コード名で登録します。</p>
          </div>
          <button type="button" onClick={onClose} disabled={isImporting} className="rounded-full p-2 text-slate-500 hover:bg-slate-100 disabled:opacity-40" aria-label="閉じる"><X size={22} /></button>
        </header>

        <div className="grid min-h-0 flex-1 grid-cols-[minmax(480px,1.15fr)_minmax(500px,1fr)] gap-4 p-4">
          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white">
            <div className="flex flex-wrap items-center gap-2 border-b border-slate-200 p-3">
              <input ref={pdfInputRef} type="file" accept="application/pdf,.pdf" className="hidden" onChange={handlePdfChange} />
              <input ref={csvInputRef} type="file" accept=".csv,text/csv" className="hidden" onChange={handleCsvChange} />
              <button type="button" onClick={() => pdfInputRef.current?.click()} disabled={isImporting} className="flex items-center gap-2 rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2 text-xs font-bold text-indigo-700 hover:bg-indigo-100 disabled:opacity-50"><FileImage size={16} />{pdfFile?.name || 'PDFを選択'}</button>
              <button type="button" onClick={() => csvInputRef.current?.click()} disabled={isImporting} className="flex items-center gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-bold text-emerald-700 hover:bg-emerald-100 disabled:opacity-50"><FileSpreadsheet size={16} />{csvFile?.name || 'CSVを選択'}</button>
              {pdfDocument && (
                <div className="ml-auto flex items-center gap-1 rounded-xl bg-slate-100 p-1">
                  <button type="button" onClick={() => setPreviewPage((page) => Math.max(1, page - 1))} disabled={previewPage <= 1 || isImporting} className="rounded-lg p-1.5 hover:bg-white disabled:opacity-30"><ChevronLeft size={16} /></button>
                  <span className="min-w-24 text-center text-xs font-bold text-slate-600">PDF {previewPage}/{pdfDocument.numPages}</span>
                  <button type="button" onClick={() => setPreviewPage((page) => Math.min(pdfDocument.numPages, page + 1))} disabled={previewPage >= pdfDocument.numPages || isImporting} className="rounded-lg p-1.5 hover:bg-white disabled:opacity-30"><ChevronRight size={16} /></button>
                </div>
              )}
            </div>

            <div className="flex items-center gap-3 border-b border-slate-200 bg-slate-50 px-3 py-2">
              <label className="flex items-center gap-2 text-xs font-bold text-slate-600">
                PDFの1ページ目＝カタログP.
                <input type="number" min="1" value={catalogStartPage} onChange={(event) => setCatalogStartPage(Math.max(1, Number(event.target.value) || 1))} disabled={isImporting} className="w-20 rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-right font-mono" />
              </label>
              <span className="text-xs text-slate-500">表示中：カタログ P.{currentCatalogPage}</span>
            </div>

            <div className="grid grid-cols-4 gap-2 border-b border-slate-200 px-3 py-2">
              {Object.entries(BOUND_LABELS).map(([key, label]) => (
                <label key={key} className="text-[10px] font-bold text-slate-500">
                  {label}（%）
                  <input type="number" min="0" max="30" step="0.1" value={bounds[key]} onChange={(event) => setBounds((current) => ({ ...current, [key]: Math.min(30, Math.max(0, Number(event.target.value) || 0)) }))} disabled={isImporting} className="mt-1 w-full rounded-lg border border-slate-300 px-2 py-1.5 text-right text-xs font-mono" />
                </label>
              ))}
            </div>

            <div className="relative min-h-0 flex-1 overflow-auto bg-slate-200 p-4">
              {!pdfDocument && <div className="flex h-full items-center justify-center text-sm font-bold text-slate-500">PDFを選択すると切り抜き範囲を確認できます</div>}
              {pdfDocument && (
                <div className="relative mx-auto w-fit max-w-full shadow-xl">
                  <canvas ref={canvasRef} className="block max-h-[620px] max-w-full bg-white object-contain" />
                  {preview.width > 0 && (
                    <div className="pointer-events-none absolute inset-0">
                      <div className="absolute border-2 border-dashed border-sky-500/80" style={{ left: `${bounds.left}%`, top: `${bounds.top}%`, right: `${bounds.right}%`, bottom: `${bounds.bottom}%` }} />
                      {previewRows.map((row) => {
                        const rect = getPdfCropRect(row, bounds);
                        const hasCode = pdfTextItemsContainCode(preview.textItems, row, bounds);
                        const hasConflict = conflictIds.has(row.id);
                        return (
                          <div key={row.id} className={`absolute border-2 ${hasConflict ? 'border-rose-500 bg-rose-400/15' : hasCode ? 'border-emerald-500 bg-emerald-400/10' : 'border-amber-500 bg-amber-400/10'}`} style={{ left: `${rect.x * 100}%`, top: `${rect.y * 100}%`, width: `${rect.width * 100}%`, height: `${rect.height * 100}%` }}>
                            <span className={`absolute left-1 top-1 rounded-md px-1.5 py-0.5 text-[10px] font-black text-white shadow ${hasConflict ? 'bg-rose-600' : hasCode ? 'bg-emerald-600' : 'bg-amber-500'}`}>{row.code || 'コード未設定'}</span>
                          </div>
                        );
                      })}
                    </div>
                  )}
                  {preview.isLoading && <div className="absolute inset-0 flex items-center justify-center bg-white/75"><Loader2 className="animate-spin text-indigo-600" size={30} /></div>}
                </div>
              )}
            </div>
          </section>

          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white">
            <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
              <div>
                <h3 className="text-sm font-black text-slate-700">切り抜き対象</h3>
                <p className="text-[11px] text-slate-500">{rows.length}件中 {selectedIds.size}件を選択</p>
              </div>
              {rows.length > 0 && <button type="button" onClick={() => setSelectedIds(selectedIds.size ? new Set() : buildSelection(rows, existingCodes))} disabled={isImporting} className="rounded-lg border border-slate-300 px-2.5 py-1.5 text-xs font-bold text-slate-600 hover:bg-slate-50">{selectedIds.size ? '全解除' : '推奨を選択'}</button>}
            </div>

            <div className="min-h-0 flex-1 overflow-auto">
              {rows.length === 0 ? (
                <div className="flex h-full items-center justify-center px-8 text-center text-sm text-slate-500">CSVを選択すると、介援隊コードとコマサイズを一覧表示します。</div>
              ) : (
                <table className="w-full border-collapse text-xs">
                  <thead className="sticky top-0 z-10 bg-slate-100 text-[10px] text-slate-500 shadow-sm">
                    <tr><th className="p-2">登録</th><th className="p-2 text-left">ページ・コード</th><th className="p-2">X/Y</th><th className="p-2 text-left">コマサイズ</th><th className="p-2">判定</th></tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => {
                      const pdfPageNumber = getPdfPageNumber(row);
                      const outsidePdf = !pdfDocument || pdfPageNumber < 1 || pdfPageNumber > pdfDocument.numPages;
                      const hasConflict = conflictIds.has(row.id);
                      const isExisting = existingCodes.has(row.code);
                      const isDuplicate = duplicateCodes.has(row.code);
                      const isCurrentPage = row.pageNumber === currentCatalogPage;
                      const hasCode = isCurrentPage && pdfTextItemsContainCode(preview.textItems, row, bounds);
                      return (
                        <tr key={row.id} className={`border-b border-slate-100 align-top ${isCurrentPage ? 'bg-indigo-50/40' : ''}`}>
                          <td className="p-2 text-center"><input type="checkbox" checked={selectedIds.has(row.id)} onChange={() => toggleSelection(row.id)} disabled={isImporting || !row.code} className="h-4 w-4 accent-indigo-600" /></td>
                          <td className="p-2">
                            <button type="button" onClick={() => !outsidePdf && setPreviewPage(pdfPageNumber)} className="mb-1 text-[10px] font-bold text-indigo-600 hover:underline">CSV P.{row.pageNumber}{outsidePdf ? '（PDF範囲外）' : ''}</button>
                            <input value={row.code} onChange={(event) => updateRow(row.id, { code: event.target.value })} onBlur={() => normalizeRowCode(row.id)} disabled={isImporting} maxLength={12} className="w-24 rounded-md border border-slate-300 px-1.5 py-1 font-mono font-bold uppercase" />
                          </td>
                          <td className="p-2">
                            <div className="flex gap-1">
                              <input type="number" min="1" max="4" value={row.xPos || ''} onChange={(event) => updateRow(row.id, { xPos: Number(event.target.value) })} disabled={isImporting} className="w-10 rounded-md border border-slate-300 px-1 py-1 text-center" aria-label={`${row.code} X座標`} />
                              <input type="number" min="1" max="4" value={row.yPos || ''} onChange={(event) => updateRow(row.id, { yPos: Number(event.target.value) })} disabled={isImporting} className="w-10 rounded-md border border-slate-300 px-1 py-1 text-center" aria-label={`${row.code} Y座標`} />
                            </div>
                            {row.positionSource === 'estimated' && <span className="mt-1 block text-[9px] font-bold text-amber-600">位置推定</span>}
                          </td>
                          <td className="p-2"><select value={row.sizeType} onChange={(event) => updateRowSize(row.id, event.target.value)} disabled={isImporting} className="max-w-40 rounded-md border border-slate-300 bg-white px-1.5 py-1">{PDF_CROP_SIZE_OPTIONS.map((option) => <option key={option} value={option}>{option}</option>)}</select></td>
                          <td className="p-2 text-center">
                            {hasConflict || outsidePdf || !row.code ? <AlertTriangle size={17} className="mx-auto text-rose-500" /> : hasCode ? <CheckCircle2 size={17} className="mx-auto text-emerald-500" /> : <AlertTriangle size={17} className="mx-auto text-amber-500" />}
                            <span className={`mt-1 block text-[9px] font-bold ${hasConflict || outsidePdf || !row.code ? 'text-rose-600' : hasCode ? 'text-emerald-600' : 'text-amber-600'}`}>{hasConflict ? '重なり' : outsidePdf ? 'ページ外' : isExisting ? '同名あり' : isDuplicate ? 'CSV重複' : hasCode ? 'コード一致' : isCurrentPage ? '要確認' : '未確認'}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {(issues.length > 0 || errorMessage) && (
              <div className="max-h-28 overflow-auto border-t border-amber-200 bg-amber-50 px-4 py-2 text-[11px] text-amber-800">
                {errorMessage && <p className="font-bold text-rose-700">{errorMessage}</p>}
                {issues.slice(0, 8).map((issue, index) => <p key={`${issue.type}-${issue.csvRow || index}`}>・{issue.message}</p>)}
                {issues.length > 8 && <p>ほか {issues.length - 8}件</p>}
              </div>
            )}

            {isImporting && (
              <div className="border-t border-indigo-100 bg-indigo-50 px-4 py-3">
                <div className="mb-1 flex justify-between text-[11px] font-bold text-indigo-700"><span>{progress.message}</span><span>{progress.current}/{progress.total}</span></div>
                <div className="h-2 overflow-hidden rounded-full bg-indigo-100"><div className="h-full bg-indigo-600 transition-all" style={{ width: `${progress.total ? (progress.current / progress.total) * 100 : 0}%` }} /></div>
              </div>
            )}

            <footer className="flex items-center justify-between border-t border-slate-200 bg-slate-50 px-4 py-3">
              <p className="text-[10px] text-slate-500">緑＝コード一致、黄＝要確認、赤＝配置修正が必要</p>
              <div className="flex gap-2">
                <button type="button" onClick={onClose} disabled={isImporting} className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-50 disabled:opacity-40">キャンセル</button>
                <button type="button" onClick={handleImport} disabled={isLocked || isImporting || !pdfDocument || importableRows.length === 0} className="flex items-center gap-2 rounded-xl bg-indigo-600 px-5 py-2 text-xs font-bold text-white shadow hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-40">{isImporting ? <Loader2 size={16} className="animate-spin" /> : <Crop size={16} />}{importableRows.length}件を切り抜いて登録</button>
              </div>
            </footer>
          </section>
        </div>
      </div>
    </div>
  );
};

export default PdfCropImportModal;
