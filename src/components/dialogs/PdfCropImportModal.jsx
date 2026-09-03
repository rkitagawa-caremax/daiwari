import { useEffect, useMemo, useRef, useState } from 'react';
import { Crop, Loader2, X } from 'lucide-react';

import {
  PdfCropFrameOverlay,
  PdfCropImportSummary,
  PdfCropManualControls,
  PdfCropPageList,
  PdfCropSourceControls
} from './PdfCropImportView';

import {
  applyPdfCropCatalogPageOverrides,
  buildPdfCropBatchPages,
  buildPdfCropPagePlans,
  DEFAULT_PDF_GRID_BOUNDS,
  extractPdfCatalogTextData,
  getPdfCropRectFromGrid,
  MAX_PDF_CROP_BATCH_PAGES,
  normalizePdfCropCode,
  parsePdfCropCsv,
  resolvePdfTextExtractionRect,
  summarizePdfCropPagePlans
} from '../../domain/pdfCropImport';
import { calibratePdfCropGrid } from '../../domain/pdfCropGridCalibration';
import { stabilizePdfCropPageRects } from '../../domain/pdfCropPageConsensus';
import {
  applyPdfCropDrag,
  clearPdfCropManualPage,
  clearPdfCropManualRect,
  countPdfCropManualRects,
  EMPTY_PDF_CROP_MANUAL_RECTS,
  clampPdfPreviewZoom,
  fitPdfPreviewSize,
  getPdfCropManualRects,
  movePdfCropRect,
  PDF_PREVIEW_ZOOM_MAX,
  PDF_PREVIEW_ZOOM_MIN,
  PDF_PREVIEW_ZOOM_STEP,
  setPdfCropManualRects,
  zoomPdfPreviewSize
} from '../../domain/pdfCropEditor';
import { mergePdfCropRects, resolvePdfCropTextRects } from '../../domain/pdfCropTextBounds';
import { readFileAutoEncoding } from '../../lib/csv';
import { cropPdfPageToFile, openPdfFile, refineCropRectToFrame, renderPdfPage } from '../../lib/pdfCropImport';

// 一括処理の流れ:
//   1. PDF を複数選択 (合計 MAX_PDF_CROP_BATCH_PAGES ページまで)。選択時に各ファイルを開いてページ数だけ読み、すぐ destroy する
//   2. ファイル名の「P010」などから対象ページ (カタログのページ番号) を自動対応させる。ページ一覧で手動変更も可能
//   3. CSV (全データ) から対象ページごとのコマを求める (除外はコード/位置不明とバッチ内の重複のみ)
//   4. 切り抜き枠は、ページ内のコードラベル座標で校正したグリッドを出発点に、描画ピクセルからコマの境界 (余白/枠線) を検出してスナップさせる
//   5. 保存時は PDF を 1 ファイルずつ開き直し、高解像度で描画 → 枠検出 → 切り抜き → canvas 解放 → destroy を繰り返す
//   6. 切り抜き済み JPEG は全ページ分まとめて onImport に渡す (App 側の登録処理は 1 回呼び出し前提のため)

const EMPTY_ROWS = Object.freeze([]);
const PREVIEW_SCALE = 2;
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
//   2. 描画ピクセルから罫線 (行=太線 / 列=点線) を探し、その内側を境界にする。罫線が無い辺は余白を使う
//   3. 検出できなかった辺だけ目印の枠まで広げ、文字が欠けないようにする
// 目印が無いコマは校正済みグリッドの推定枠を出発点にする (ラベルが多いページほど探索範囲を狭める)。
const MIN_ANCHORS_FOR_TIGHT_SEARCH = 3;
const ANCHORED_SEARCH_TOLERANCE = 0.08;
const CALIBRATED_SEARCH_TOLERANCE = 0.12;
const UNCALIBRATED_SEARCH_TOLERANCE = 0.25;

// 矢印キーでの移動量 (ページ比)。Shift でざっくり動かす
const NUDGE_STEP = 0.002;
const NUDGE_STEP_LARGE = 0.01;

const resolveCropRect = (canvas, row, grid, textRects, manualRect) => {
  // 手で決めた枠があれば自動補正より優先する
  if (manualRect) return { rect: manualRect, snappedEdges: {}, snappedCount: 0, hasTextAnchor: false, isManual: true };
  const textRect = textRects?.get(row.id) || null;
  const base = textRect || getPdfCropRectFromGrid(row, grid);
  const searchToleranceRatio = textRect
    ? ANCHORED_SEARCH_TOLERANCE
    : (grid.anchorCount >= MIN_ANCHORS_FOR_TIGHT_SEARCH ? CALIBRATED_SEARCH_TOLERANCE : UNCALIBRATED_SEARCH_TOLERANCE);
  const refined = refineCropRectToFrame(canvas, base, { searchToleranceRatio });
  return {
    rect: mergePdfCropRects(refined.rect, textRect, refined.snappedEdges),
    snappedEdges: refined.snappedEdges,
    snappedCount: refined.snappedCount,
    hasTextAnchor: !!textRect,
    isManual: false
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
  const stageRef = useRef(null);
  const frameLayerRef = useRef(null);
  const dragRef = useRef(null);
  const [pdfSources, setPdfSources] = useState([]);
  const [catalogPageOverrides, setCatalogPageOverrides] = useState({});
  const [activeBatchPageId, setActiveBatchPageId] = useState('');
  const [csvFile, setCsvFile] = useState(null);
  const [pdfDocument, setPdfDocument] = useState(null);
  const [allRows, setAllRows] = useState([]);
  const [issues, setIssues] = useState([]);
  const [preview, setPreview] = useState({ isLoading: false, textItems: [], width: 0, height: 0, version: 0 });
  const [errorMessage, setErrorMessage] = useState('');
  const [isReadingPdfs, setIsReadingPdfs] = useState(false);
  const [skipExistingCodes, setSkipExistingCodes] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, message: '' });
  // 手で調整した切り抜き枠 { [ページid]: { [コマid]: rect } }
  const [manualRects, setManualRects] = useState(EMPTY_PDF_CROP_MANUAL_RECTS);
  const [selectedRowIds, setSelectedRowIds] = useState(() => new Set());
  const [isDraggingFrame, setIsDraggingFrame] = useState(false);
  const [stageSize, setStageSize] = useState({ width: 0, height: 0 });
  const [zoom, setZoom] = useState(PDF_PREVIEW_ZOOM_MIN);

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
    calibratePdfCropGrid({ rows: activeTargetRows, textItems: preview.textItems, bounds: DEFAULT_PDF_GRID_BOUNDS })
  ), [activeTargetRows, preview.textItems]);

  // コマ左上の番号ラベル / コードラベルを目印にした切り抜き枠 (見つかったコマのみ)
  const previewTextRects = useMemo(() => (
    resolvePdfCropTextRects({ rows: activeTargetRows, textItems: preview.textItems, grid: previewGrid })
  ), [activeTargetRows, preview.textItems, previewGrid]);

  const activePageManualRects = getPdfCropManualRects(manualRects, activeBatchPage?.id);

  // 自動で求めた枠。ピクセル解析を含むのでドラッグ中に作り直さないよう、手動ぶんとは分けて memo する
  const autoPreviewRects = useMemo(() => {
    const canvas = canvasRef.current;
    const entries = activeTargetRows.map((row) => {
      const base = previewTextRects.get(row.id) || getPdfCropRectFromGrid(row, previewGrid);
      if (!canvas || !preview.width || preview.isLoading) return { row, rect: base, snappedEdges: {}, snappedCount: 0, hasTextAnchor: previewTextRects.has(row.id), isManual: false };
      return { row, ...resolveCropRect(canvas, row, previewGrid, previewTextRects) };
    });
    return stabilizePdfCropPageRects({ entries, grid: previewGrid });
  }, [activeTargetRows, preview, previewGrid, previewTextRects]);

  // プレビュー上の切り抜き枠 (手動で決めた枠が最優先)
  const previewRects = useMemo(() => autoPreviewRects.map((entry) => {
    const manualRect = activePageManualRects[entry.row.id];
    return manualRect ? { ...entry, rect: manualRect, snappedCount: 0, hasTextAnchor: false, isManual: true } : entry;
  }), [activePageManualRects, autoPreviewRects]);
  const previewSnappedCount = previewRects.filter((entry) => entry.snappedCount >= 2 || entry.hasTextAnchor || entry.isManual).length;
  const selectedEntries = previewRects.filter((entry) => selectedRowIds.has(entry.row.id));
  const selectedManualCount = selectedEntries.filter((entry) => entry.isManual).length;
  const pageManualCount = Object.keys(activePageManualRects).length;
  const totalManualCount = countPdfCropManualRects(manualRects);

  // 表示領域いっぱいにページを収めるサイズ (枠の座標はこの表示サイズを基準に px → 比率へ換算する)
  useEffect(() => {
    const node = stageRef.current;
    if (!node || typeof ResizeObserver === 'undefined') return undefined;
    const observer = new ResizeObserver(([entry]) => {
      const box = entry?.contentRect;
      if (box) setStageSize({ width: box.width, height: box.height });
    });
    observer.observe(node);
    return () => observer.disconnect();
  }, []);
  const stageFit = useMemo(() => fitPdfPreviewSize(preview, stageSize), [preview, stageSize]);
  const stageDisplay = useMemo(() => zoomPdfPreviewSize(stageFit, zoom), [stageFit, zoom]);

  // ページを切り替えたら選択を外す (枠は手動調整ぶんを残したまま)
  useEffect(() => setSelectedRowIds(new Set()), [activeBatchPage?.id]);

  // --- コマ枠の手動調整 ---
  // Ctrl (Mac は ⌘) + クリックで複数選択でき、移動・変形・矢印キーは選択中の全枠へ同時にかかる。
  const startFrameDrag = (event, entry, handle) => {
    const layer = frameLayerRef.current;
    if (!layer || isImporting) return;
    const bounds = layer.getBoundingClientRect();
    if (!bounds.width || !bounds.height) return;
    if (handle !== 'move') {
      // つまみを押したときはフォーカスを枠に残す (矢印キーでの微調整を続けられるように)
      event.preventDefault();
      event.stopPropagation();
    }

    // Ctrl+クリックは選択のトグルだけ行い、ドラッグは開始しない
    if (handle === 'move' && (event.ctrlKey || event.metaKey)) {
      setSelectedRowIds((current) => {
        const next = new Set(current);
        if (next.has(entry.row.id)) next.delete(entry.row.id);
        else next.add(entry.row.id);
        return next;
      });
      return;
    }

    // 選択中の枠を掴んだらグループごと、未選択の枠を掴んだらその枠だけの選択にする
    const nextSelected = selectedRowIds.has(entry.row.id) ? selectedRowIds : new Set([entry.row.id]);
    if (nextSelected !== selectedRowIds) setSelectedRowIds(nextSelected);
    const items = previewRects
      .filter((candidate) => nextSelected.has(candidate.row.id))
      .map((candidate) => ({ rowId: candidate.row.id, rect: candidate.rect }));

    dragRef.current = {
      pointerId: event.pointerId,
      pageId: activeBatchPage?.id,
      handle,
      startX: event.clientX,
      startY: event.clientY,
      items,
      width: bounds.width,
      height: bounds.height
    };
    setIsDraggingFrame(true);
  };

  useEffect(() => {
    if (!isDraggingFrame) return undefined;
    const handleMove = (event) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      const dx = (event.clientX - drag.startX) / drag.width;
      const dy = (event.clientY - drag.startY) / drag.height;
      const rectsById = Object.fromEntries(drag.items.map(({ rowId, rect }) => (
        [rowId, applyPdfCropDrag(rect, drag.handle, dx, dy)]
      )));
      setManualRects((current) => setPdfCropManualRects(current, drag.pageId, rectsById));
    };
    const handleEnd = () => {
      dragRef.current = null;
      setIsDraggingFrame(false);
    };
    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleEnd);
    window.addEventListener('pointercancel', handleEnd);
    return () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleEnd);
      window.removeEventListener('pointercancel', handleEnd);
    };
  }, [isDraggingFrame]);

  const handleFrameKeyDown = (event, entry) => {
    if (event.key === 'Escape') {
      setSelectedRowIds(new Set());
      return;
    }
    const step = event.shiftKey ? NUDGE_STEP_LARGE : NUDGE_STEP;
    const nudge = {
      ArrowLeft: [-step, 0],
      ArrowRight: [step, 0],
      ArrowUp: [0, -step],
      ArrowDown: [0, step]
    }[event.key];
    if (!nudge || isImporting) return;
    event.preventDefault();
    const isGrouped = selectedRowIds.has(entry.row.id);
    const targets = isGrouped
      ? previewRects.filter((candidate) => selectedRowIds.has(candidate.row.id))
      : [entry];
    if (!isGrouped) setSelectedRowIds(new Set([entry.row.id]));
    const rectsById = Object.fromEntries(targets.map((candidate) => (
      [candidate.row.id, movePdfCropRect(candidate.rect, nudge[0], nudge[1])]
    )));
    setManualRects((current) => setPdfCropManualRects(current, activeBatchPage?.id, rectsById));
  };

  const resetSelectedFrames = () => {
    if (selectedRowIds.size === 0) return;
    setManualRects((current) => (
      [...selectedRowIds].reduce((acc, rowId) => clearPdfCropManualRect(acc, activeBatchPage?.id, rowId), current)
    ));
  };

  const resetPageFrames = () => setManualRects((current) => clearPdfCropManualPage(current, activeBatchPage?.id));

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
        setErrorMessage(`プレビューを表示できません。${error?.message || ''}`);
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
        setErrorMessage(`一度に扱えるのは${MAX_PDF_CROP_BATCH_PAGES}ページまでです（選んだのは${sources.length}ファイル・${totalPages}ページ）`);
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
      setManualRects(EMPTY_PDF_CROP_MANUAL_RECTS);
      if (unreadable.length > 0) setErrorMessage(`読み込めなかったPDFは除きました: ${unreadable.join(', ')}`);
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
      setManualRects(EMPTY_PDF_CROP_MANUAL_RECTS);
      if (parsed.rows.length === 0) setErrorMessage('このCSVに切り抜けるコマがありません');
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
            setProgress({ current: completed, total, message: `${page.filename}（P.${page.catalogPage}）を読み込み中…` });
            const rendered = await renderPdfPage(sourceDocument, page.pdfPageNumber, { scale: EXPORT_SCALE });
            try {
              // このページのグリッドを高解像度描画の文字レイヤーで校正し、目印の座標も取り直してから切り抜く
              const pageGrid = calibratePdfCropGrid({ rows: plan.targetRows, textItems: rendered.textItems, bounds: DEFAULT_PDF_GRID_BOUNDS });
              const pageTextRects = resolvePdfCropTextRects({ rows: plan.targetRows, textItems: rendered.textItems, grid: pageGrid });
              const pageManualRects = getPdfCropManualRects(manualRects, page.id);
              const pageAutoEntries = stabilizePdfCropPageRects({
                grid: pageGrid,
                entries: plan.targetRows.map((row) => ({
                  row,
                  ...resolveCropRect(rendered.canvas, row, pageGrid, pageTextRects)
                }))
              });
              const pageAutoRects = new Map(pageAutoEntries.map((entry) => [entry.row.id, entry.rect]));
              for (const row of importRows) {
                const code = normalizePdfCropCode(row.code);
                const manualRect = pageManualRects[row.id];
                const cropRect = manualRect || pageAutoRects.get(row.id) || getPdfCropRectFromGrid(row, pageGrid);
                const textExtractionRect = resolvePdfTextExtractionRect({
                  cropRect,
                  gridRect: getPdfCropRectFromGrid(row, pageGrid),
                  textRect: pageTextRects.get(row.id),
                  isManual: !!manualRect
                });
                const sourceTextData = extractPdfCatalogTextData({
                  textItems: rendered.textItems,
                  rect: textExtractionRect,
                  code,
                  catalogName: row.catalogName
                });
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
                  textExtractionRect,
                  ...sourceTextData
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
      setErrorMessage(`画像を保存できませんでした。${error?.message || ''}`);
    } finally {
      setIsImporting(false);
    }
  };

  const hasPdf = batchPages.length > 0;
  const isBusy = isImporting || isReadingPdfs;
  const isReady = hasPdf && !!csvFile && !isBusy && batchSummary.importCount > 0;
  const warnings = [
    hasPdf && csvFile && batchSummary.importCount === 0 ? '保存できるコマがありません（対象ページを確認）' : '',
    csvFile && batchSummary.pagesWithoutCsv > 0 ? `CSVに無いページのPDFが${batchSummary.pagesWithoutCsv}件（P.を選び直す）` : '',
    batchSummary.duplicateCatalogPages > 0 ? '同じP.に複数のPDFが割り当て済み' : '',
    batchSummary.unplaceableCount > 0 ? `コードか位置が無い${batchSummary.unplaceableCount}件は除外` : '',
    batchSummary.duplicateCodeCount > 0 ? `重複したコード${batchSummary.duplicateCodeCount}件は除外` : '',
    issues.length > 0 ? `CSVの${issues.length}行は対象外（保存に影響なし）` : ''
  ].filter(Boolean);

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/55 p-1 backdrop-blur-sm">
      <div className="flex h-[99vh] w-[min(1480px,98vw)] flex-col overflow-hidden rounded-2xl border border-white/30 bg-slate-100 shadow-2xl" role="dialog" aria-modal="true" aria-label="PDF画像取り込み">
        <header className="flex items-center justify-between gap-3 border-b border-slate-200 bg-white px-4 py-1.5">
          <h2 className="flex min-w-0 items-center gap-2 text-sm font-black text-slate-800" title={`校正PDF（最大${MAX_PDF_CROP_BATCH_PAGES}ページ）と全データCSVを選ぶと、商品コマを介援隊コード名の画像として保存します。`}>
            <Crop size={16} className="shrink-0 text-indigo-600" />PDF画像取り込み
          </h2>
          <button type="button" onClick={onClose} disabled={isImporting} className="rounded-full p-1.5 text-slate-500 hover:bg-slate-100 disabled:opacity-40" aria-label="閉じる"><X size={18} /></button>
        </header>

        <div className="grid min-h-0 flex-1 grid-cols-[minmax(0,1fr)_300px] gap-2 p-2">
          <section className="flex min-h-0 flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white">
            <div className="flex items-center gap-2 border-b border-slate-200 px-3 py-2">
              <p className="min-w-0 flex-1 truncate text-xs font-bold text-slate-700">
                {activeBatchPage ? <>{activeBatchPage.filename} <span className="text-slate-400">→</span> P.{activeBatchPage.catalogPage}{csvFile && <span className="ml-2 font-medium text-slate-500">{activeTargetRows.length}コマ</span>}</> : 'プレビュー'}
              </p>
              {selectedEntries.length > 0 && (
                <span className="shrink-0 rounded-full bg-indigo-50 px-2.5 py-1 text-[10px] font-bold text-indigo-700">選択中 {selectedEntries.length === 1 ? normalizePdfCropCode(selectedEntries[0].row.code) : `${selectedEntries.length}コマ`}</span>
              )}
              {selectedManualCount > 0 && (
                <button type="button" onClick={resetSelectedFrames} disabled={isImporting} className="shrink-0 rounded-full border border-indigo-200 px-2.5 py-1 text-[10px] font-bold text-indigo-700 hover:bg-indigo-50 disabled:opacity-40">{selectedEntries.length === 1 ? 'この枠を自動に戻す' : `選択${selectedManualCount}件を自動に戻す`}</button>
              )}
              {activeTargetRows.length > 0 && preview.width > 0 && (
                <span className="shrink-0 rounded-full bg-slate-100 px-2.5 py-1 text-[10px] font-bold text-slate-600" title="自動＝罫線や余白から境界を決めた枠。目印＝コマ番号やコードを見つけた枠。手動＝自分で調整した枠。（境界はページ全体で突き合わせて外れ値を除いています）">枠 自動{previewSnappedCount}/{activeTargetRows.length}・目印{previewTextRects.size}{pageManualCount > 0 ? `・手動${pageManualCount}` : ''}</span>
              )}
            </div>
            {/* ページを表示領域いっぱいに収め、その上に切り抜き枠を重ねる (枠はドラッグで調整できる) */}
            <div ref={stageRef} className="relative min-h-0 flex-1 overflow-hidden bg-slate-300 p-1">
              <div className="flex h-full w-full overflow-auto">
                {!hasPdf && <p className="m-auto text-sm font-bold text-slate-500">PDFを選ぶとページが表示されます</p>}
                {hasPdf && (
                  <div
                    className="relative m-auto shadow-xl"
                    style={stageDisplay.width > 0 ? { width: `${stageDisplay.width}px`, height: `${stageDisplay.height}px` } : { width: '100%', height: '100%' }}
                  >
                    <canvas ref={canvasRef} className="block h-full w-full bg-white" />
                    <div
                      ref={frameLayerRef}
                      className="absolute inset-0"
                      onPointerDown={(event) => { if (event.target === event.currentTarget) setSelectedRowIds(new Set()); }}
                    >
                      {preview.width > 0 && previewRects.map((entry) => (
                        <PdfCropFrameOverlay
                          key={entry.row.id}
                          entry={entry}
                          textItems={preview.textItems}
                          isSelected={selectedRowIds.has(entry.row.id)}
                          onPointerDown={startFrameDrag}
                          onKeyDown={handleFrameKeyDown}
                        />
                      ))}
                    </div>
                    {(preview.isLoading || !pdfDocument) && <div className="absolute inset-0 z-30 flex items-center justify-center bg-white/75"><Loader2 className="animate-spin text-indigo-600" size={30} /></div>}
                  </div>
                )}
              </div>
              {hasPdf && (
                <div className="absolute bottom-2 right-2 z-40 flex items-center gap-1.5 rounded-full border border-slate-200 bg-white/95 py-1 pl-2 pr-1 shadow-lg backdrop-blur">
                  <button type="button" onClick={() => setZoom((current) => clampPdfPreviewZoom(current - PDF_PREVIEW_ZOOM_STEP))} disabled={zoom <= PDF_PREVIEW_ZOOM_MIN} className="rounded-full px-1.5 text-sm font-black text-slate-600 hover:bg-slate-100 disabled:opacity-30" aria-label="縮小">−</button>
                  <input
                    type="range"
                    min={PDF_PREVIEW_ZOOM_MIN}
                    max={PDF_PREVIEW_ZOOM_MAX}
                    step="0.05"
                    value={zoom}
                    onChange={(event) => setZoom(clampPdfPreviewZoom(event.target.value))}
                    className="w-28 accent-indigo-600"
                    aria-label="ページの表示倍率"
                  />
                  <button type="button" onClick={() => setZoom((current) => clampPdfPreviewZoom(current + PDF_PREVIEW_ZOOM_STEP))} disabled={zoom >= PDF_PREVIEW_ZOOM_MAX} className="rounded-full px-1.5 text-sm font-black text-slate-600 hover:bg-slate-100 disabled:opacity-30" aria-label="拡大">＋</button>
                  <button type="button" onClick={() => setZoom(PDF_PREVIEW_ZOOM_MIN)} className="w-14 rounded-full bg-slate-100 py-0.5 text-[10px] font-black text-slate-600 hover:bg-slate-200" title="クリックで全体表示に戻す">{Math.round(zoom * 100)}%</button>
                </div>
              )}
            </div>
          </section>

          <aside className="flex min-h-0 flex-col gap-1.5">
            <PdfCropSourceControls
              hasPdf={hasPdf}
              isReadingPdfs={isReadingPdfs}
              isBusy={isBusy}
              pdfFileCount={pdfSources.length}
              pdfPageCount={batchPages.length}
              maxPages={MAX_PDF_CROP_BATCH_PAGES}
              csvFile={csvFile}
              pdfInputRef={pdfInputRef}
              csvInputRef={csvInputRef}
              onPdfChange={handlePdfChange}
              onCsvChange={handleCsvChange}
            />

            <PdfCropImportSummary
              summary={batchSummary}
              skipExistingCodes={skipExistingCodes}
              onSkipExistingCodesChange={(event) => setSkipExistingCodes(event.target.checked)}
              isImporting={isImporting}
              errorMessage={errorMessage}
              warnings={warnings}
            />

            <PdfCropManualControls
              hasPdf={hasPdf}
              totalManualCount={totalManualCount}
              pageManualCount={pageManualCount}
              isImporting={isImporting}
              onResetPageFrames={resetPageFrames}
            />

            <PdfCropPageList
              pagePlans={pagePlans}
              csvFile={csvFile}
              activePageId={activeBatchPage?.id}
              availablePages={availablePages}
              pdfSources={pdfSources}
              manualRects={manualRects}
              isImporting={isImporting}
              onSelectPage={setActiveBatchPageId}
              onCatalogPageChange={handleCatalogPageChange}
            />

          </aside>
        </div>

        {isImporting && (
          <div className="border-t border-indigo-100 bg-indigo-50 px-6 py-3">
            <div className="mb-1 flex justify-between text-[11px] font-bold text-indigo-700"><span>{progress.message}</span><span>{progress.current}/{progress.total}</span></div>
            <div className="h-2 overflow-hidden rounded-full bg-indigo-100"><div className="h-full bg-indigo-600 transition-all" style={{ width: `${progress.total ? (progress.current / progress.total) * 100 : 0}%` }} /></div>
          </div>
        )}

        <footer className="flex items-center justify-end gap-2 border-t border-slate-200 bg-white px-4 py-1.5">
          <button type="button" onClick={onClose} disabled={isImporting} className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-50 disabled:opacity-40">キャンセル</button>
          <button type="button" onClick={handleImport} disabled={isLocked || !isReady} className="flex items-center gap-2 rounded-xl bg-indigo-600 px-6 py-2 text-xs font-bold text-white shadow hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-40">{isImporting ? <Loader2 size={16} className="animate-spin" /> : <Crop size={16} />}{batchSummary.importCount}コマを保存</button>
        </footer>
      </div>
    </div>
  );
};

export default PdfCropImportModal;
