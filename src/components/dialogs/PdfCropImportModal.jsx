import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Crop,
  Database,
  FileImage,
  FileSpreadsheet,
  Loader2,
  X
} from 'lucide-react';

import {
  applyPdfCropCatalogPageOverrides,
  buildPdfCropBatchPages,
  buildPdfCropPagePlans,
  DEFAULT_PDF_GRID_BOUNDS,
  extractPdfCropText,
  getPdfCropRect,
  MAX_PDF_CROP_BATCH_PAGES,
  normalizePdfCropCode,
  parsePdfCropCsv,
  pdfTextItemsContainCode,
  summarizePdfCropPagePlans
} from '../../domain/pdfCropImport';
import { readFileAutoEncoding } from '../../lib/csv';
import { cropPdfPageToFile, openPdfFile, renderPdfPage } from '../../lib/pdfCropImport';

// 一括処理の流れ:
//   1. PDF を複数選択 (合計 MAX_PDF_CROP_BATCH_PAGES ページまで)。選択時に各ファイルを開いてページ数だけ読み、すぐ destroy する
//   2. ファイル名の「P010」などから対象ページ (カタログのページ番号) を自動対応させる。手動で変更も可能
//   3. CSV (全データ) から対象ページごとのコマを求め、既存画像・重複コード・座標衝突を除外する
//   4. 保存時は PDF を 1 ファイルずつ開き直し、ページを高解像度で描画 → 切り抜き → canvas 解放 → destroy を繰り返す
//      (描画キャンバスは 1 ページあたり十数 MB あるため、複数ページを同時に保持しない)
//   5. 切り抜き済み JPEG は全ページ分まとめて onImport に渡す (App 側の登録処理は 1 回呼び出し前提のため)

const BOUND_LABELS = Object.freeze({
  left: '左',
  top: '上',
  right: '右',
  bottom: '下'
});

const EMPTY_ROWS = Object.freeze([]);
const EMPTY_ID_SET = new Set();
const PREVIEW_SCALE = 1.4;
const EXPORT_SCALE = 3;

const getExistingCodeSet = (images = []) => new Set(images
  .map((image) => normalizePdfCropCode(image.code || image.name || ''))
  .filter(Boolean));

const isPdfFile = (file) => file && (file.type === 'application/pdf' || /\.pdf$/i.test(file.name || ''));

const sortFilesByName = (files) => [...files].sort((left, right) => (
  String(left.name || '').localeCompare(String(right.name || ''), 'ja', { numeric: true })
));

const yieldToUi = () => new Promise((resolve) => setTimeout(resolve, 0));

const releaseCanvas = (canvas) => {
  if (!canvas) return;
  canvas.width = 1;
  canvas.height = 1;
};

const destroyDocument = async (pdfDocument) => {
  try {
    await pdfDocument?.destroy?.();
  } catch (error) {
    console.warn('PDF document cleanup failed:', error);
  }
};

const PdfCropImportModal = ({ isOpen, onClose, onImport, existingImages = [], isLocked = false }) => {
  const canvasRef = useRef(null);
  const pdfInputRef = useRef(null);
  const csvInputRef = useRef(null);
  const documentRef = useRef(null);
  const documentFileRef = useRef(null);
  const renderSequenceRef = useRef(0);
  const documentLoadSequenceRef = useRef(0);
  const [pdfSources, setPdfSources] = useState([]);
  const [catalogPageOverrides, setCatalogPageOverrides] = useState({});
  const [activeBatchPageId, setActiveBatchPageId] = useState('');
  const [csvFile, setCsvFile] = useState(null);
  const [pdfDocument, setPdfDocument] = useState(null);
  const [allRows, setAllRows] = useState([]);
  const [issues, setIssues] = useState([]);
  const [preview, setPreview] = useState({ isLoading: false, textItems: [], width: 0, height: 0 });
  const [bounds, setBounds] = useState({ ...DEFAULT_PDF_GRID_BOUNDS });
  const [errorMessage, setErrorMessage] = useState('');
  const [isReadingPdfs, setIsReadingPdfs] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, message: '' });

  const existingCodes = useMemo(() => getExistingCodeSet(existingImages), [existingImages]);
  const batchPages = useMemo(() => (
    applyPdfCropCatalogPageOverrides(buildPdfCropBatchPages(pdfSources), catalogPageOverrides)
  ), [catalogPageOverrides, pdfSources]);
  const pagePlans = useMemo(() => (
    buildPdfCropPagePlans({ batchPages, rows: allRows, existingCodes })
  ), [allRows, batchPages, existingCodes]);
  const batchSummary = useMemo(() => summarizePdfCropPagePlans(pagePlans), [pagePlans]);
  const activePlan = useMemo(() => (
    pagePlans.find((plan) => plan.page.id === activeBatchPageId) || pagePlans[0] || null
  ), [activeBatchPageId, pagePlans]);
  const activeBatchPage = activePlan?.page || null;
  const activeCatalogPage = activeBatchPage?.catalogPage || 1;
  const activePdfPageNumber = activeBatchPage?.pdfPageNumber || 1;
  const activeTargetRows = activePlan?.targetRows || EMPTY_ROWS;
  const activeConflictIds = activePlan?.conflictIds || EMPTY_ID_SET;
  const availablePages = useMemo(() => [...new Set(allRows.map((row) => row.pageNumber))].sort((a, b) => a - b), [allRows]);
  const catalogPageOptions = useMemo(() => (
    availablePages.includes(activeCatalogPage) ? availablePages : [activeCatalogPage, ...availablePages]
  ), [activeCatalogPage, availablePages]);

  const codeMatchCount = useMemo(() => activeTargetRows.filter((row) => (
    pdfTextItemsContainCode(preview.textItems, row, bounds)
  )).length, [activeTargetRows, bounds, preview.textItems]);
  const textDetectedCount = useMemo(() => activeTargetRows.filter((row) => (
    extractPdfCropText(preview.textItems, row, bounds).text.length > 0
  )).length, [activeTargetRows, bounds, preview.textItems]);

  useEffect(() => () => {
    renderSequenceRef.current++;
    documentLoadSequenceRef.current++;
    const current = documentRef.current;
    documentRef.current = null;
    documentFileRef.current = null;
    destroyDocument(current);
  }, []);

  // プレビュー対象のファイルが変わったら PDF を開き直す (同じファイル内のページ切替は描画 effect だけで済む)
  useEffect(() => {
    if (!isOpen) return undefined;
    const file = activeBatchPage?.file || null;
    if (!file) {
      const current = documentRef.current;
      documentRef.current = null;
      documentFileRef.current = null;
      if (current) {
        setPdfDocument(null);
        destroyDocument(current);
      }
      return undefined;
    }
    if (documentFileRef.current === file && documentRef.current) return undefined;

    const sequence = ++documentLoadSequenceRef.current;
    let cancelled = false;
    setPreview((current) => ({ ...current, isLoading: true }));

    openPdfFile(file)
      .then(async (nextDocument) => {
        if (cancelled || sequence !== documentLoadSequenceRef.current) {
          await destroyDocument(nextDocument);
          return;
        }
        const previous = documentRef.current;
        documentRef.current = nextDocument;
        documentFileRef.current = file;
        setPdfDocument(nextDocument);
        await destroyDocument(previous);
      })
      .catch((error) => {
        if (cancelled || sequence !== documentLoadSequenceRef.current) return;
        console.error('PDF loading failed:', error);
        setPreview((current) => ({ ...current, isLoading: false }));
        setErrorMessage(`PDFを読み込めません。${error?.message || ''}`);
      });

    return () => {
      cancelled = true;
    };
  }, [activeBatchPage, isOpen]);

  useEffect(() => {
    if (!isOpen || !pdfDocument || !canvasRef.current) return undefined;
    const sequence = ++renderSequenceRef.current;
    let cancelled = false;
    setPreview((current) => ({ ...current, isLoading: true }));

    renderPdfPage(pdfDocument, activePdfPageNumber, { scale: PREVIEW_SCALE, canvas: canvasRef.current })
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
  }, [activePdfPageNumber, isOpen, pdfDocument]);

  if (!isOpen) return null;

  const handlePdfChange = async (event) => {
    const files = sortFilesByName(Array.from(event.target.files || []).filter(isPdfFile));
    event.target.value = '';
    if (files.length === 0) return;
    setErrorMessage('');
    setIsReadingPdfs(true);
    try {
      const sources = [];
      const unreadable = [];
      for (const file of files) {
        let probe = null;
        try {
          probe = await openPdfFile(file);
          sources.push({ file, numPages: probe.numPages || 1 });
        } catch (error) {
          console.error('PDF probing failed:', file.name, error);
          unreadable.push(file.name);
        } finally {
          await destroyDocument(probe);
        }
      }

      const totalPages = sources.reduce((sum, source) => sum + source.numPages, 0);
      if (totalPages > MAX_PDF_CROP_BATCH_PAGES) {
        setErrorMessage(`一度に処理できるのは合計${MAX_PDF_CROP_BATCH_PAGES}ページまでです（選択: ${sources.length}ファイル / ${totalPages}ページ）。ファイル数を減らしてください。`);
        return;
      }
      if (sources.length === 0) {
        setErrorMessage(`PDFを読み込めません。${unreadable.join(', ')}`);
        return;
      }

      // 以前のプレビュー用ドキュメントを解放してから差し替える
      const previous = documentRef.current;
      documentRef.current = null;
      documentFileRef.current = null;
      setPdfDocument(null);
      setPreview({ isLoading: false, textItems: [], width: 0, height: 0 });
      await destroyDocument(previous);

      setPdfSources(sources);
      setCatalogPageOverrides({});
      setActiveBatchPageId('');
      if (unreadable.length > 0) {
        setErrorMessage(`読み込めなかったPDFは除外しました: ${unreadable.join(', ')}`);
      }
    } finally {
      setIsReadingPdfs(false);
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
      setAllRows(parsed.rows);
      setIssues(parsed.issues);
      if (parsed.rows.length === 0) setErrorMessage('切り抜き対象の介援隊コードがCSVにありません。');
    } catch (error) {
      console.error('Crop CSV loading failed:', error);
      setErrorMessage(`CSVを読み込めません。${error?.message || ''}`);
    }
  };

  const handleCatalogPageChange = (batchPageId, value) => {
    const nextPage = Number.parseInt(value, 10);
    if (!Number.isInteger(nextPage) || nextPage < 1) return;
    setCatalogPageOverrides((current) => ({ ...current, [batchPageId]: nextPage }));
  };

  const handleImport = async () => {
    if (isLocked || isImporting || isReadingPdfs || batchSummary.importCount === 0) return;
    setIsImporting(true);
    setErrorMessage('');
    const total = batchSummary.importCount;
    let completed = 0;
    setProgress({ current: 0, total, message: 'PDFを1ページずつ切り抜いています…' });

    try {
      const items = [];
      // 同じファイルのページはまとめて 1 回だけ開く
      const plansBySource = new Map();
      pagePlans.forEach((plan) => {
        if (plan.importRows.length === 0) return;
        const key = plan.page.sourceIndex;
        if (!plansBySource.has(key)) plansBySource.set(key, []);
        plansBySource.get(key).push(plan);
      });

      for (const plans of plansBySource.values()) {
        const sourceDocument = await openPdfFile(plans[0].page.file);
        try {
          for (const plan of plans) {
            const { page, importRows } = plan;
            setProgress({ current: completed, total, message: `${page.filename}（P.${page.catalogPage}）を高解像度で描画しています…` });
            const rendered = await renderPdfPage(sourceDocument, page.pdfPageNumber, { scale: EXPORT_SCALE });
            try {
              for (const row of importRows) {
                const code = normalizePdfCropCode(row.code);
                const cropRect = getPdfCropRect(row, bounds);
                const sourceTextData = extractPdfCropText(rendered.textItems, row, bounds);
                const file = await cropPdfPageToFile({
                  canvas: rendered.canvas,
                  normalizedRect: cropRect,
                  filename: `${code}.jpg`
                });
                items.push({
                  file,
                  code,
                  sourcePdfName: page.filename,
                  sourcePage: page.catalogPage,
                  pdfPageNumber: page.pdfPageNumber,
                  sizeType: row.sizeType,
                  cropRect,
                  sourceText: sourceTextData.text,
                  sourceTextTruncated: sourceTextData.truncated
                });
                completed++;
                setProgress({ current: completed, total, message: `P.${page.catalogPage} ${code} を切り抜きました（${completed}/${total}）` });
              }
            } finally {
              // 高解像度キャンバスはページごとに即解放する
              releaseCanvas(rendered.canvas);
            }
            await yieldToUi();
          }
        } finally {
          await destroyDocument(sourceDocument);
        }
      }

      const result = await onImport(items, ({ current, total: uploadTotal, message }) => {
        setProgress({ current, total: uploadTotal, message });
      });
      if (!result || result.successCount > 0) onClose();
    } catch (error) {
      console.error('PDF crop import failed:', error);
      setErrorMessage(`切り抜き画像の登録に失敗しました。${error?.message || ''}`);
    } finally {
      setIsImporting(false);
    }
  };

  const hasPdf = batchPages.length > 0;
  const isReady = hasPdf && !!csvFile && !isReadingPdfs && batchSummary.importCount > 0;
  const isBusy = isImporting || isReadingPdfs;
  const pdfSummaryLabel = hasPdf
    ? `${pdfSources.length}ファイル / ${batchPages.length}ページ（最大${MAX_PDF_CROP_BATCH_PAGES}ページ）`
    : `P010.pdf のような1ページPDFを最大${MAX_PDF_CROP_BATCH_PAGES}ファイル`;

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/55 p-4 backdrop-blur-sm">
      <div className="flex h-[min(900px,94vh)] w-[min(1180px,96vw)] flex-col overflow-hidden rounded-3xl border border-white/30 bg-slate-100 shadow-2xl" role="dialog" aria-modal="true" aria-label="PDF画像取り込み">
        <header className="flex items-center justify-between border-b border-slate-200 bg-white px-6 py-4">
          <div>
            <h2 className="flex items-center gap-2 text-lg font-black text-slate-800"><Crop size={20} className="text-indigo-600" />PDF画像取り込み</h2>
            <p className="mt-1 text-xs text-slate-500">校正PDF（最大{MAX_PDF_CROP_BATCH_PAGES}ページ）と全データCSVを選ぶと、ファイル名のP番号から対象ページを自動で対応させ、商品コマを画像ライブラリへ保存します。</p>
          </div>
          <button type="button" onClick={onClose} disabled={isImporting} className="rounded-full p-2 text-slate-500 hover:bg-slate-100 disabled:opacity-40" aria-label="閉じる"><X size={22} /></button>
        </header>

        <div className="grid grid-cols-3 gap-3 border-b border-slate-200 bg-white px-5 py-4">
          <button type="button" onClick={() => pdfInputRef.current?.click()} disabled={isBusy} className={`flex min-w-0 items-center gap-3 rounded-2xl border p-3 text-left transition-colors disabled:opacity-50 ${hasPdf ? 'border-indigo-300 bg-indigo-50' : 'border-slate-200 hover:bg-slate-50'}`}>
            <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-sm font-black ${hasPdf ? 'bg-indigo-600 text-white' : 'bg-slate-200 text-slate-600'}`}>1</span>
            {isReadingPdfs ? <Loader2 size={19} className="shrink-0 animate-spin text-indigo-600" /> : <FileImage size={19} className="shrink-0 text-indigo-600" />}
            <span className="min-w-0"><span className="block text-xs font-black text-slate-700">PDFを選択（複数可）</span><span className="block truncate text-[10px] text-slate-500">{isReadingPdfs ? 'PDFを確認しています…' : pdfSummaryLabel}</span></span>
          </button>
          <button type="button" onClick={() => csvInputRef.current?.click()} disabled={isBusy} className={`flex min-w-0 items-center gap-3 rounded-2xl border p-3 text-left transition-colors disabled:opacity-50 ${csvFile ? 'border-emerald-300 bg-emerald-50' : 'border-slate-200 hover:bg-slate-50'}`}>
            <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-sm font-black ${csvFile ? 'bg-emerald-600 text-white' : 'bg-slate-200 text-slate-600'}`}>2</span>
            <FileSpreadsheet size={19} className="shrink-0 text-emerald-600" />
            <span className="min-w-0"><span className="block text-xs font-black text-slate-700">CSV（全データ）を選択</span><span className="block truncate text-[10px] text-slate-500">{csvFile?.name || '台割CSV'}</span></span>
          </button>
          <div className={`flex min-w-0 items-center gap-3 rounded-2xl border p-3 ${activeTargetRows.length ? 'border-sky-300 bg-sky-50' : 'border-slate-200'}`}>
            <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-sm font-black ${activeTargetRows.length ? 'bg-sky-600 text-white' : 'bg-slate-200 text-slate-600'}`}>3</span>
            <label className="min-w-0 flex-1 text-xs font-black text-slate-700">プレビュー
              <select value={activeBatchPage?.id || ''} onChange={(event) => setActiveBatchPageId(event.target.value)} disabled={!hasPdf || isImporting} className="mt-1 w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-sm font-bold disabled:bg-slate-100">
                {!hasPdf && <option value="">PDF選択後に指定</option>}
                {batchPages.map((page) => (
                  <option key={page.id} value={page.id}>
                    {page.filename}{pdfSources[page.sourceIndex]?.numPages > 1 ? ` (${page.pdfPageNumber})` : ''} → P.{page.catalogPage}
                  </option>
                ))}
              </select>
            </label>
            <label className="min-w-24 text-[10px] font-bold text-slate-500">対象ページ
              <select value={activeCatalogPage} onChange={(event) => activeBatchPage && handleCatalogPageChange(activeBatchPage.id, event.target.value)} disabled={!hasPdf || !availablePages.length || isImporting} className="mt-1 w-full rounded-lg border border-slate-300 bg-white px-1.5 py-1.5 text-xs font-bold disabled:bg-slate-100">
                {!availablePages.length && <option value={activeCatalogPage}>CSV選択後</option>}
                {availablePages.length > 0 && catalogPageOptions.map((page) => <option key={page} value={page}>P.{page}</option>)}
              </select>
            </label>
          </div>
          <input ref={pdfInputRef} type="file" accept="application/pdf,.pdf" multiple className="hidden" onChange={handlePdfChange} />
          <input ref={csvInputRef} type="file" accept=".csv,text/csv" className="hidden" onChange={handleCsvChange} />
        </div>

        <div className="grid min-h-0 flex-1 grid-cols-[minmax(0,1fr)_320px] gap-4 p-4">
          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white">
            <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
              <div><h3 className="text-sm font-black text-slate-700">切り抜きプレビュー</h3><p className="text-[10px] text-slate-500">{activeBatchPage ? `${activeBatchPage.filename} → P.${activeCatalogPage} の${activeTargetRows.length}コマを表示しています` : 'PDFを選択すると表示されます'}</p></div>
              {hasPdf && <span className="rounded-full bg-slate-100 px-3 py-1 text-[10px] font-bold text-slate-600">ページ {Math.max(1, batchPages.findIndex((page) => page.id === activeBatchPage?.id) + 1)}/{batchPages.length}</span>}
            </div>
            <div className="relative min-h-0 flex-1 overflow-auto bg-slate-200 p-4">
              {!hasPdf && <div className="flex h-full items-center justify-center text-sm font-bold text-slate-500">最初にPDFを選択してください</div>}
              {hasPdf && (
                <div className="relative mx-auto w-fit max-w-full shadow-xl">
                  <canvas ref={canvasRef} className="block max-h-[600px] max-w-full bg-white object-contain" />
                  {preview.width > 0 && (
                    <div className="pointer-events-none absolute inset-0">
                      {activeTargetRows.map((row) => {
                        const rect = getPdfCropRect(row, bounds);
                        const hasCode = pdfTextItemsContainCode(preview.textItems, row, bounds);
                        const hasConflict = activeConflictIds.has(row.id);
                        return (
                          <div key={row.id} className={`absolute border-2 ${hasConflict ? 'border-rose-500 bg-rose-400/15' : hasCode ? 'border-emerald-500 bg-emerald-400/10' : 'border-amber-500 bg-amber-400/10'}`} style={{ left: `${rect.x * 100}%`, top: `${rect.y * 100}%`, width: `${rect.width * 100}%`, height: `${rect.height * 100}%` }}>
                            <span className={`absolute left-1 top-1 rounded-md px-1.5 py-0.5 text-[10px] font-black text-white shadow ${hasConflict ? 'bg-rose-600' : hasCode ? 'bg-emerald-600' : 'bg-amber-500'}`}>{row.code}</span>
                          </div>
                        );
                      })}
                    </div>
                  )}
                  {(preview.isLoading || (!pdfDocument && hasPdf)) && <div className="absolute inset-0 flex items-center justify-center bg-white/75"><Loader2 className="animate-spin text-indigo-600" size={30} /></div>}
                </div>
              )}
            </div>
            <details className="border-t border-slate-200 bg-slate-50 px-4 py-2">
              <summary className="cursor-pointer text-[10px] font-bold text-slate-500">切り抜き枠がずれる場合だけ調整（全ページ共通）</summary>
              <div className="mt-2 grid grid-cols-4 gap-2">
                {Object.entries(BOUND_LABELS).map(([key, label]) => (
                  <label key={key} className="text-[10px] font-bold text-slate-500">{label}余白（%）
                    <input type="number" min="0" max="30" step="0.1" value={bounds[key]} onChange={(event) => setBounds((current) => ({ ...current, [key]: Math.min(30, Math.max(0, Number(event.target.value) || 0)) }))} disabled={isImporting} className="mt-1 w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-right font-mono text-xs" />
                  </label>
                ))}
              </div>
            </details>
          </section>

          <aside className="flex min-h-0 flex-col gap-3">
            <section className="rounded-2xl border border-slate-200 bg-white p-4">
              <h3 className="text-sm font-black text-slate-700">取り込み内容（{batchSummary.pageCount}ページ）</h3>
              <div className="mt-3 grid grid-cols-2 gap-2">
                <div className="rounded-xl bg-indigo-50 p-3"><span className="block text-[10px] font-bold text-indigo-500">対象コマ</span><span className="text-xl font-black text-indigo-700">{batchSummary.targetCount}</span></div>
                <div className="rounded-xl bg-emerald-50 p-3"><span className="block text-[10px] font-bold text-emerald-600">新規保存</span><span className="text-xl font-black text-emerald-700">{batchSummary.importCount}</span></div>
                <div className="rounded-xl bg-sky-50 p-3"><span className="block text-[10px] font-bold text-sky-600">コード一致<span className="ml-1 font-medium text-sky-400">表示中</span></span><span className="text-xl font-black text-sky-700">{codeMatchCount}</span></div>
                <div className="rounded-xl bg-violet-50 p-3"><span className="block text-[10px] font-bold text-violet-600">文字保存<span className="ml-1 font-medium text-violet-400">表示中</span></span><span className="text-xl font-black text-violet-700">{textDetectedCount}</span></div>
              </div>
              {batchSummary.skippedCount > 0 && <p className="mt-3 text-[10px] font-bold text-amber-700">既存画像・重複・配置エラーの{batchSummary.skippedCount}件は自動で除外します。</p>}
            </section>

            <section className="min-h-0 flex-1 overflow-auto rounded-2xl border border-slate-200 bg-white p-3">
              <h3 className="px-1 text-xs font-black text-slate-700">ページ別の保存内容</h3>
              {pagePlans.length === 0 && <p className="mt-3 px-1 text-xs text-slate-500">PDFとCSVを選択すると表示されます。</p>}
              <ul className="mt-2 space-y-1.5">
                {pagePlans.map((plan) => {
                  const isActive = plan.page.id === activeBatchPage?.id;
                  const statusClass = !csvFile
                    ? 'text-slate-400'
                    : !plan.hasCsvRows
                      ? 'text-rose-600'
                      : plan.importRows.length === 0
                        ? 'text-amber-600'
                        : 'text-emerald-600';
                  const statusLabel = !csvFile
                    ? 'CSV待ち'
                    : !plan.hasCsvRows
                      ? 'CSVに該当ページなし'
                      : `${plan.importRows.length}/${plan.targetRows.length}コマ`;
                  return (
                    <li key={plan.page.id}>
                      <button
                        type="button"
                        onClick={() => setActiveBatchPageId(plan.page.id)}
                        disabled={isImporting}
                        className={`w-full rounded-xl border px-2.5 py-2 text-left transition-colors ${isActive ? 'border-indigo-300 bg-indigo-50' : 'border-slate-200 hover:bg-slate-50'}`}
                      >
                        <span className="flex items-center justify-between gap-2">
                          <span className="min-w-0 truncate text-[11px] font-bold text-slate-700" title={plan.page.filename}>{plan.page.filename}</span>
                          <span className="shrink-0 rounded-md bg-white px-1.5 py-0.5 font-mono text-[10px] font-black text-slate-600 ring-1 ring-slate-200">P.{plan.page.catalogPage}</span>
                        </span>
                        <span className={`mt-0.5 block text-[10px] font-bold ${statusClass}`}>
                          {statusLabel}
                          {plan.isDuplicateCatalogPage && <span className="ml-1 text-rose-600">・対象ページが重複</span>}
                          {plan.conflictIds.size > 0 && <span className="ml-1 text-rose-600">・重なり{plan.conflictIds.size}件</span>}
                        </span>
                        {plan.importRows.length > 0 && (
                          <span className="mt-1 flex flex-wrap gap-1">
                            {plan.importRows.map((row) => <span key={row.id} className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-[9px] font-bold text-slate-600">{normalizePdfCropCode(row.code)}.jpg</span>)}
                          </span>
                        )}
                      </button>
                    </li>
                  );
                })}
              </ul>
            </section>

            <section className="rounded-2xl border border-violet-200 bg-violet-50 p-3 text-[10px] leading-relaxed text-violet-800">
              <div className="flex gap-2"><Database size={16} className="shrink-0" /><p><strong>PDF内の文字も保存します</strong><br />各コマ内の文字を最大4,000文字に整形し、画面には表示せず画像データへ保持します。将来の検索・突合機能で利用できます。</p></div>
            </section>

            {(errorMessage || issues.length > 0 || batchSummary.conflictCount > 0 || batchSummary.duplicateCatalogPages > 0 || (csvFile && batchSummary.pagesWithoutCsv > 0)) && (
              <section className="max-h-24 overflow-auto rounded-xl border border-amber-200 bg-amber-50 p-3 text-[10px] text-amber-800">
                {errorMessage && <p className="font-bold text-rose-700">{errorMessage}</p>}
                {csvFile && batchSummary.pagesWithoutCsv > 0 && <p>・CSVに該当ページがないPDFが{batchSummary.pagesWithoutCsv}件あります。「対象ページ」で指定してください。</p>}
                {batchSummary.duplicateCatalogPages > 0 && <p>・同じ対象ページに割り当てられたPDFがあります。後のPDFは重複コードとして除外されます。</p>}
                {batchSummary.conflictCount > 0 && <p>・座標が重なっているコマが{batchSummary.conflictCount}件あります。</p>}
                {issues.length > 0 && <p>・CSV全体で読み取れない行が{issues.length}件あります。</p>}
              </section>
            )}
          </aside>
        </div>

        {isImporting && (
          <div className="border-t border-indigo-100 bg-indigo-50 px-6 py-3">
            <div className="mb-1 flex justify-between text-[11px] font-bold text-indigo-700"><span>{progress.message}</span><span>{progress.current}/{progress.total}</span></div>
            <div className="h-2 overflow-hidden rounded-full bg-indigo-100"><div className="h-full bg-indigo-600 transition-all" style={{ width: `${progress.total ? (progress.current / progress.total) * 100 : 0}%` }} /></div>
          </div>
        )}

        <footer className="flex items-center justify-between border-t border-slate-200 bg-white px-6 py-4">
          <p className="flex items-center gap-1.5 text-[10px] text-slate-500">{isReady ? <CheckCircle2 size={15} className="text-emerald-500" /> : <AlertTriangle size={15} className="text-amber-500" />}{isReady ? '取り込み準備ができました' : 'PDF・CSV・対象ページを確認してください'}</p>
          <div className="flex gap-2">
            <button type="button" onClick={onClose} disabled={isImporting} className="rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-xs font-bold text-slate-600 hover:bg-slate-50 disabled:opacity-40">キャンセル</button>
            <button type="button" onClick={handleImport} disabled={isLocked || isBusy || !isReady} className="flex items-center gap-2 rounded-xl bg-indigo-600 px-6 py-2.5 text-xs font-bold text-white shadow hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-40">{isImporting ? <Loader2 size={16} className="animate-spin" /> : <Crop size={16} />}{batchSummary.pageCount}ページ・{batchSummary.importCount}コマを保存</button>
          </div>
        </footer>
      </div>
    </div>
  );
};

export default PdfCropImportModal;
