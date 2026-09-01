import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Crop, FileImage, FileSpreadsheet, Loader2, X } from 'lucide-react';

import {
  applyPdfCropCatalogPageOverrides,
  buildPdfCropBatchPages,
  buildPdfCropPagePlans,
  DEFAULT_PDF_GRID_BOUNDS,
  extractPdfTextInRect,
  getPdfCropRectFromGrid,
  MAX_PDF_CROP_BATCH_PAGES,
  normalizePdfCropCode,
  parsePdfCropCsv,
  pdfTextItemsContainCodeInRect,
  summarizePdfCropPagePlans
} from '../../domain/pdfCropImport';
import { calibratePdfCropGrid } from '../../domain/pdfCropGridCalibration';
import { resolvePdfCropTextRects, unionPdfCropRects } from '../../domain/pdfCropTextBounds';
import { readFileAutoEncoding } from '../../lib/csv';
import { cropPdfPageToFile, openPdfFile, refineCropRectToFrame, renderPdfPage } from '../../lib/pdfCropImport';

// 一括処理の流れ:
//   1. PDF を複数選択 (合計 MAX_PDF_CROP_BATCH_PAGES ページまで)。選択時に各ファイルを開いてページ数だけ読み、すぐ destroy する
//   2. ファイル名の「P010」などから対象ページ (カタログのページ番号) を自動対応させる。ページ一覧で手動変更も可能
//   3. CSV (全データ) から対象ページごとのコマを求める (除外はコード/位置不明とバッチ内の重複のみ)
//   4. 切り抜き枠は、ページ内のコードラベル座標で校正したグリッドを出発点に、描画ピクセルからコマの境界 (余白/枠線) を検出してスナップさせる
//   5. 保存時は PDF を 1 ファイルずつ開き直し、高解像度で描画 → 枠検出 → 切り抜き → canvas 解放 → destroy を繰り返す
//   6. 切り抜き済み JPEG は全ページ分まとめて onImport に渡す (App 側の登録処理は 1 回呼び出し前提のため)

const BOUND_LABELS = Object.freeze({ left: '左', top: '上', right: '右', bottom: '下' });
const EMPTY_ROWS = Object.freeze([]);
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

// 切り抜き枠の決め方 (プレビューと保存で同じ関数を使う):
//   1. コマ左上の番号ラベルと右下のメーカー名などの目印から求めた枠 (textRect) があればそれを出発点にする
//   2. 描画ピクセルから余白/枠線を探して外側の境界へ広げる
//   3. 目印の枠と和を取り、文字が欠けないようにする
// 目印が無いコマは校正済みグリッドの推定枠を出発点にする (ラベルが多いページほど探索範囲を狭める)。
const MIN_ANCHORS_FOR_TIGHT_SEARCH = 3;
const ANCHORED_SEARCH_TOLERANCE = 0.08;
const CALIBRATED_SEARCH_TOLERANCE = 0.12;
const UNCALIBRATED_SEARCH_TOLERANCE = 0.25;

const resolveCropRect = (canvas, row, grid, textRects) => {
  const textRect = textRects?.get(row.id) || null;
  const base = textRect || getPdfCropRectFromGrid(row, grid);
  const searchToleranceRatio = textRect
    ? ANCHORED_SEARCH_TOLERANCE
    : (grid.anchorCount >= MIN_ANCHORS_FOR_TIGHT_SEARCH ? CALIBRATED_SEARCH_TOLERANCE : UNCALIBRATED_SEARCH_TOLERANCE);
  const refined = refineCropRectToFrame(canvas, base, { searchToleranceRatio });
  return {
    rect: unionPdfCropRects(refined.rect, textRect),
    snappedCount: refined.snappedCount,
    hasTextAnchor: !!textRect
  };
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
  const [preview, setPreview] = useState({ isLoading: false, textItems: [], width: 0, height: 0, version: 0 });
  const [bounds, setBounds] = useState({ ...DEFAULT_PDF_GRID_BOUNDS });
  const [errorMessage, setErrorMessage] = useState('');
  const [isReadingPdfs, setIsReadingPdfs] = useState(false);
  const [skipExistingCodes, setSkipExistingCodes] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, message: '' });

  const existingCodes = useMemo(() => getExistingCodeSet(existingImages), [existingImages]);
  const batchPages = useMemo(() => (
    applyPdfCropCatalogPageOverrides(buildPdfCropBatchPages(pdfSources), catalogPageOverrides)
  ), [catalogPageOverrides, pdfSources]);
  const pagePlans = useMemo(() => (
    buildPdfCropPagePlans({ batchPages, rows: allRows, existingCodes, skipExistingCodes })
  ), [allRows, batchPages, existingCodes, skipExistingCodes]);
  const batchSummary = useMemo(() => summarizePdfCropPagePlans(pagePlans), [pagePlans]);
  const activePlan = useMemo(() => (
    pagePlans.find((plan) => plan.page.id === activeBatchPageId) || pagePlans[0] || null
  ), [activeBatchPageId, pagePlans]);
  const activeBatchPage = activePlan?.page || null;
  const activePdfPageNumber = activeBatchPage?.pdfPageNumber || 1;
  const activeTargetRows = activePlan?.targetRows || EMPTY_ROWS;
  const availablePages = useMemo(() => [...new Set(allRows.map((row) => row.pageNumber))].sort((a, b) => a - b), [allRows]);

  // プレビュー中ページのグリッド校正 (文字レイヤーのコードラベル座標から)
  const previewGrid = useMemo(() => (
    calibratePdfCropGrid({ rows: activeTargetRows, textItems: preview.textItems, bounds })
  ), [activeTargetRows, bounds, preview.textItems]);

  // コマ左上の番号ラベル / コードラベルを目印にした切り抜き枠 (見つかったコマのみ)
  const previewTextRects = useMemo(() => (
    resolvePdfCropTextRects({ rows: activeTargetRows, textItems: preview.textItems, grid: previewGrid })
  ), [activeTargetRows, preview.textItems, previewGrid]);

  // プレビュー上の切り抜き枠 (境界スナップ済み)。preview.version は描画完了ごとに進む
  const previewRects = useMemo(() => {
    const canvas = canvasRef.current;
    return activeTargetRows.map((row) => {
      const base = previewTextRects.get(row.id) || getPdfCropRectFromGrid(row, previewGrid);
      if (!canvas || !preview.width || preview.isLoading) return { row, rect: base, snappedCount: 0, hasTextAnchor: previewTextRects.has(row.id) };
      return { row, ...resolveCropRect(canvas, row, previewGrid, previewTextRects) };
    });
  }, [activeTargetRows, preview, previewGrid, previewTextRects]);
  const previewSnappedCount = previewRects.filter((entry) => entry.snappedCount >= 2 || entry.hasTextAnchor).length;

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
        setPreview((current) => ({
          isLoading: false,
          textItems: result.textItems,
          width: result.width,
          height: result.height,
          version: current.version + 1
        }));
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
        setErrorMessage(`一度に処理できるのは合計${MAX_PDF_CROP_BATCH_PAGES}ページまでです（選択: ${sources.length}ファイル / ${totalPages}ページ）。`);
        return;
      }
      if (sources.length === 0) {
        setErrorMessage(`PDFを読み込めません。${unreadable.join(', ')}`);
        return;
      }

      const previous = documentRef.current;
      documentRef.current = null;
      documentFileRef.current = null;
      setPdfDocument(null);
      setPreview((current) => ({ isLoading: false, textItems: [], width: 0, height: 0, version: current.version + 1 }));
      await destroyDocument(previous);

      setPdfSources(sources);
      setCatalogPageOverrides({});
      setActiveBatchPageId('');
      if (unreadable.length > 0) setErrorMessage(`読み込めなかったPDFは除外しました: ${unreadable.join(', ')}`);
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
            setProgress({ current: completed, total, message: `${page.filename}（P.${page.catalogPage}）を描画しています…` });
            const rendered = await renderPdfPage(sourceDocument, page.pdfPageNumber, { scale: EXPORT_SCALE });
            try {
              // このページのグリッドを高解像度描画の文字レイヤーで校正し、目印の座標も取り直してから切り抜く
              const pageGrid = calibratePdfCropGrid({ rows: plan.targetRows, textItems: rendered.textItems, bounds });
              const pageTextRects = resolvePdfCropTextRects({ rows: plan.targetRows, textItems: rendered.textItems, grid: pageGrid });
              for (const row of importRows) {
                const code = normalizePdfCropCode(row.code);
                const { rect: cropRect } = resolveCropRect(rendered.canvas, row, pageGrid, pageTextRects);
                const sourceTextData = extractPdfTextInRect(rendered.textItems, cropRect);
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
  const isBusy = isImporting || isReadingPdfs;
  const isReady = hasPdf && !!csvFile && !isBusy && batchSummary.importCount > 0;
  const statusText = !hasPdf
    ? 'PDFを選択してください'
    : !csvFile
      ? 'CSV（全データ）を選択してください'
      : batchSummary.importCount === 0
        ? '保存できるコマがありません。ページ一覧で対象ページを確認してください'
        : `${batchSummary.pageCount}ページ・${batchSummary.importCount}コマを保存できます`;
  const warnings = [
    csvFile && batchSummary.pagesWithoutCsv > 0 ? `CSVに該当ページがないPDFが${batchSummary.pagesWithoutCsv}件（ページ一覧で対象ページを変更）` : '',
    batchSummary.duplicateCatalogPages > 0 ? '同じ対象ページに複数のPDFが割り当てられています' : '',
    batchSummary.unplaceableCount > 0 ? `コード・位置が読めない${batchSummary.unplaceableCount}件は除外` : '',
    batchSummary.duplicateCodeCount > 0 ? `同じコードの重複${batchSummary.duplicateCodeCount}件は除外` : '',
    issues.length > 0 ? `CSVで読み取れない行が${issues.length}件（保存には影響しません）` : ''
  ].filter(Boolean);

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/55 p-4 backdrop-blur-sm">
      <div className="flex h-[min(860px,94vh)] w-[min(1120px,96vw)] flex-col overflow-hidden rounded-3xl border border-white/30 bg-slate-100 shadow-2xl" role="dialog" aria-modal="true" aria-label="PDF画像取り込み">
        <header className="flex items-center justify-between border-b border-slate-200 bg-white px-6 py-4">
          <div>
            <h2 className="flex items-center gap-2 text-lg font-black text-slate-800"><Crop size={20} className="text-indigo-600" />PDF画像取り込み</h2>
            <p className="mt-1 text-xs text-slate-500">校正PDF（最大{MAX_PDF_CROP_BATCH_PAGES}ページ）と全データCSVを選ぶと、商品コマを介援隊コード名の画像として保存します。</p>
          </div>
          <button type="button" onClick={onClose} disabled={isImporting} className="rounded-full p-2 text-slate-500 hover:bg-slate-100 disabled:opacity-40" aria-label="閉じる"><X size={22} /></button>
        </header>

        <div className="grid grid-cols-2 gap-3 border-b border-slate-200 bg-white px-5 py-3">
          <button type="button" onClick={() => pdfInputRef.current?.click()} disabled={isBusy} className={`flex min-w-0 items-center gap-3 rounded-2xl border px-4 py-3 text-left transition-colors disabled:opacity-50 ${hasPdf ? 'border-indigo-300 bg-indigo-50' : 'border-dashed border-slate-300 hover:bg-slate-50'}`}>
            {isReadingPdfs ? <Loader2 size={20} className="shrink-0 animate-spin text-indigo-600" /> : <FileImage size={20} className="shrink-0 text-indigo-600" />}
            <span className="min-w-0">
              <span className="block text-sm font-black text-slate-700">{hasPdf ? `PDF ${pdfSources.length}ファイル / ${batchPages.length}ページ` : 'PDFを選択（複数可）'}</span>
              <span className="block truncate text-[11px] text-slate-500">{isReadingPdfs ? 'PDFを確認しています…' : hasPdf ? 'クリックで選び直し' : `P010.pdf のような校正PDFを最大${MAX_PDF_CROP_BATCH_PAGES}ページ`}</span>
            </span>
          </button>
          <button type="button" onClick={() => csvInputRef.current?.click()} disabled={isBusy} className={`flex min-w-0 items-center gap-3 rounded-2xl border px-4 py-3 text-left transition-colors disabled:opacity-50 ${csvFile ? 'border-emerald-300 bg-emerald-50' : 'border-dashed border-slate-300 hover:bg-slate-50'}`}>
            <FileSpreadsheet size={20} className="shrink-0 text-emerald-600" />
            <span className="min-w-0">
              <span className="block text-sm font-black text-slate-700">{csvFile ? 'CSV 選択済み' : 'CSV（全データ）を選択'}</span>
              <span className="block truncate text-[11px] text-slate-500">{csvFile?.name || '台割の出力CSV'}</span>
            </span>
          </button>
          <input ref={pdfInputRef} type="file" accept="application/pdf,.pdf" multiple className="hidden" onChange={handlePdfChange} />
          <input ref={csvInputRef} type="file" accept=".csv,text/csv" className="hidden" onChange={handleCsvChange} />
        </div>

        <div className="grid min-h-0 flex-1 grid-cols-[minmax(0,1fr)_300px] gap-4 p-4">
          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white">
            <div className="flex items-center justify-between gap-3 border-b border-slate-200 px-4 py-2.5">
              <p className="min-w-0 truncate text-xs font-bold text-slate-700">
                {activeBatchPage ? <>{activeBatchPage.filename} <span className="text-slate-400">→</span> P.{activeBatchPage.catalogPage}{csvFile && <span className="ml-2 font-medium text-slate-500">{activeTargetRows.length}コマ</span>}</> : 'プレビュー'}
              </p>
              {activeTargetRows.length > 0 && preview.width > 0 && (
                <span className="shrink-0 rounded-full bg-slate-100 px-2.5 py-1 text-[10px] font-bold text-slate-600" title="コマ左上の番号ラベルとコードラベルを目印に切り抜き枠を合わせ、さらに余白/枠線を検出して外側の境界へ広げたコマ数">枠を自動補正 {previewSnappedCount}/{activeTargetRows.length}・目印{previewTextRects.size}</span>
              )}
            </div>
            <div className="relative min-h-0 flex-1 overflow-auto bg-slate-200 p-4">
              {!hasPdf && <div className="flex h-full items-center justify-center text-sm font-bold text-slate-500">PDFを選択するとプレビューが表示されます</div>}
              {hasPdf && (
                <div className="relative mx-auto w-fit max-w-full shadow-xl">
                  <canvas ref={canvasRef} className="block max-h-[620px] max-w-full bg-white object-contain" />
                  {preview.width > 0 && (
                    <div className="pointer-events-none absolute inset-0">
                      {previewRects.map(({ row, rect }) => {
                        const code = normalizePdfCropCode(row.code);
                        const hasCode = pdfTextItemsContainCodeInRect(preview.textItems, code, rect);
                        return (
                          <div key={row.id} className={`absolute border-2 ${hasCode ? 'border-emerald-500 bg-emerald-400/10' : 'border-amber-500 bg-amber-400/10'}`} style={{ left: `${rect.x * 100}%`, top: `${rect.y * 100}%`, width: `${rect.width * 100}%`, height: `${rect.height * 100}%` }}>
                            <span className={`absolute left-1 top-1 rounded-md px-1.5 py-0.5 text-[10px] font-black text-white shadow ${hasCode ? 'bg-emerald-600' : 'bg-amber-500'}`}>{code}</span>
                          </div>
                        );
                      })}
                    </div>
                  )}
                  {(preview.isLoading || !pdfDocument) && <div className="absolute inset-0 flex items-center justify-center bg-white/75"><Loader2 className="animate-spin text-indigo-600" size={30} /></div>}
                </div>
              )}
            </div>
            <details className="border-t border-slate-200 bg-slate-50 px-4 py-2">
              <summary className="cursor-pointer text-[10px] font-bold text-slate-500">枠が大きくずれる場合だけ余白を調整（全ページ共通）</summary>
              <div className="mt-2 grid grid-cols-4 gap-2">
                {Object.entries(BOUND_LABELS).map(([key, label]) => (
                  <label key={key} className="text-[10px] font-bold text-slate-500">{label}（%）
                    <input type="number" min="0" max="30" step="0.1" value={bounds[key]} onChange={(event) => setBounds((current) => ({ ...current, [key]: Math.min(30, Math.max(0, Number(event.target.value) || 0)) }))} disabled={isImporting} className="mt-1 w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-right font-mono text-xs" />
                  </label>
                ))}
              </div>
            </details>
          </section>

          <aside className="flex min-h-0 flex-col gap-3">
            <section className="rounded-2xl border border-slate-200 bg-white p-4">
              <p className="text-[10px] font-bold text-slate-500">保存するコマ</p>
              <p className="mt-0.5 flex items-baseline gap-1.5"><span className="text-3xl font-black text-indigo-700">{batchSummary.importCount}</span><span className="text-xs font-bold text-slate-500">コマ / {batchSummary.pageCount}ページ</span></p>
              {batchSummary.existingCount > 0 && (
                <label className="mt-2 flex items-start gap-2 text-[10px] font-bold text-slate-600">
                  <input type="checkbox" checked={skipExistingCodes} onChange={(event) => setSkipExistingCodes(event.target.checked)} disabled={isImporting} className="mt-0.5" />
                  <span>ライブラリに同じコードがある{batchSummary.existingCount}件をスキップする（未チェックなら追加登録）</span>
                </label>
              )}
              {(errorMessage || warnings.length > 0) && (
                <div className="mt-2 space-y-0.5 text-[10px] text-amber-800">
                  {errorMessage && <p className="font-bold text-rose-700">{errorMessage}</p>}
                  {warnings.map((warning) => <p key={warning}>・{warning}</p>)}
                </div>
              )}
            </section>

            <section className="min-h-0 flex-1 overflow-auto rounded-2xl border border-slate-200 bg-white p-2">
              <p className="px-2 pb-1 pt-1 text-[10px] font-bold text-slate-500">ページ一覧（クリックでプレビュー / 対象ページは変更可）</p>
              {pagePlans.length === 0 && <p className="px-2 py-3 text-xs text-slate-500">PDFを選択すると表示されます。</p>}
              <ul className="space-y-1">
                {pagePlans.map((plan) => {
                  const isActive = plan.page.id === activeBatchPage?.id;
                  const pageOptions = availablePages.includes(plan.page.catalogPage) ? availablePages : [plan.page.catalogPage, ...availablePages];
                  const countLabel = !csvFile ? '' : !plan.hasCsvRows ? 'CSVに該当なし' : `${plan.importRows.length}コマ`;
                  const countClass = !csvFile ? 'text-slate-400' : !plan.hasCsvRows ? 'text-rose-600' : plan.importRows.length === 0 ? 'text-amber-600' : 'text-emerald-600';
                  return (
                    <li key={plan.page.id}>
                      <div
                        role="button"
                        tabIndex={0}
                        onClick={() => setActiveBatchPageId(plan.page.id)}
                        onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setActiveBatchPageId(plan.page.id); } }}
                        className={`flex items-center gap-2 rounded-xl border px-2.5 py-2 transition-colors ${isActive ? 'border-indigo-300 bg-indigo-50' : 'border-slate-200 hover:bg-slate-50'}`}
                      >
                        <span className="min-w-0 flex-1 truncate text-[11px] font-bold text-slate-700" title={plan.page.filename}>{plan.page.filename}{pdfSources[plan.page.sourceIndex]?.numPages > 1 ? ` (${plan.page.pdfPageNumber})` : ''}</span>
                        <select
                          value={plan.page.catalogPage}
                          onClick={(event) => event.stopPropagation()}
                          onChange={(event) => handleCatalogPageChange(plan.page.id, event.target.value)}
                          disabled={isImporting}
                          className="shrink-0 rounded-md border border-slate-300 bg-white px-1 py-0.5 font-mono text-[10px] font-bold text-slate-700"
                          aria-label={`${plan.page.filename} の対象ページ`}
                        >
                          {pageOptions.map((page) => <option key={page} value={page}>P.{page}</option>)}
                        </select>
                        <span className={`w-16 shrink-0 text-right text-[10px] font-bold ${countClass}`}>{countLabel}</span>
                      </div>
                    </li>
                  );
                })}
              </ul>
            </section>
          </aside>
        </div>

        {isImporting && (
          <div className="border-t border-indigo-100 bg-indigo-50 px-6 py-3">
            <div className="mb-1 flex justify-between text-[11px] font-bold text-indigo-700"><span>{progress.message}</span><span>{progress.current}/{progress.total}</span></div>
            <div className="h-2 overflow-hidden rounded-full bg-indigo-100"><div className="h-full bg-indigo-600 transition-all" style={{ width: `${progress.total ? (progress.current / progress.total) * 100 : 0}%` }} /></div>
          </div>
        )}

        <footer className="flex items-center justify-between border-t border-slate-200 bg-white px-6 py-4">
          <p className={`text-[11px] font-bold ${isReady ? 'text-emerald-600' : 'text-slate-500'}`}>{statusText}</p>
          <div className="flex gap-2">
            <button type="button" onClick={onClose} disabled={isImporting} className="rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-xs font-bold text-slate-600 hover:bg-slate-50 disabled:opacity-40">キャンセル</button>
            <button type="button" onClick={handleImport} disabled={isLocked || !isReady} className="flex items-center gap-2 rounded-xl bg-indigo-600 px-6 py-2.5 text-xs font-bold text-white shadow hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-40">{isImporting ? <Loader2 size={16} className="animate-spin" /> : <Crop size={16} />}{batchSummary.importCount}コマを保存</button>
          </div>
        </footer>
      </div>
    </div>
  );
};

export default PdfCropImportModal;
