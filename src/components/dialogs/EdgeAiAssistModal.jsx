import { useCallback, useMemo, useRef, useState } from 'react';
import {
  ArrowRight,
  CheckCircle2,
  Database,
  FileDiff,
  FileUp,
  Image as ImageIcon,
  Loader2,
  MapPin,
  Search,
  ShieldCheck,
  Sparkles,
  X
} from 'lucide-react';

import {
  EDGE_AI_MODEL_VERSION,
  buildEdgeCatalogProducts,
  compareCatalogSnapshots,
  createEdgeCatalogFingerprint,
  parseCatalogSnapshotCsv,
  rankEdgeCatalogProducts,
  rankSimilarEdgeCatalogProducts,
  summarizeCatalogDiff
} from '../../domain/edgeAiCatalog';
import { embedEdgeAiTexts, initializeEdgeAi } from '../../lib/edgeAiClient';
import { readFileAutoEncoding } from '../../lib/csv';
import { idbHelper } from '../../idbHelper';

const INDEX_CACHE_KEY = `edgeAiCatalogIndex:${EDGE_AI_MODEL_VERSION}`;
const EMBEDDING_BATCH_SIZE = 12;

const TABS = [
  { id: 'search', label: '商品意味検索' },
  { id: 'similar', label: '類似品提案' },
  { id: 'diff', label: 'CSV差分' }
];

const STATUS_LABELS = {
  added: '追加',
  removed: '削除',
  modified: '変更',
  unchanged: '変更なし'
};

const STATUS_STYLES = {
  added: 'bg-emerald-50 text-emerald-700 border-emerald-200',
  removed: 'bg-rose-50 text-rose-700 border-rose-200',
  modified: 'bg-amber-50 text-amber-700 border-amber-200',
  unchanged: 'bg-slate-50 text-slate-500 border-slate-200'
};

const formatProgress = (progress) => {
  if (!progress) return 'モデルを準備しています…';
  if (progress.status === 'fallback') return '互換モードへ切り替えています…';
  const percent = Number(progress.progress);
  const filename = String(progress.file || progress.name || '').split('/').pop();
  if (Number.isFinite(percent)) return `${filename || 'モデル'} ${Math.round(percent)}%`;
  const labels = {
    initiate: '読込準備中',
    download: 'モデル読込中',
    progress: 'モデル読込中',
    done: '読込完了',
    ready: '起動中'
  };
  return labels[progress.status] || 'モデルを準備しています…';
};

const percent = (score) => `${Math.round(Math.max(0, Math.min(1, score || 0)) * 100)}%`;

const ResultCard = ({ result, onSelectSimilar, onOpenSheet }) => {
  const { product, score } = result;
  const assignment = product.assignments?.[0];
  return (
    <article className="group grid grid-cols-[72px_minmax(0,1fr)] gap-4 rounded-[22px] bg-white px-4 py-3.5 shadow-[0_1px_3px_rgba(15,23,42,0.08)] transition hover:shadow-[0_5px_20px_rgba(15,23,42,0.09)] sm:grid-cols-[84px_minmax(0,1fr)]">
      <div className="flex h-[72px] items-center justify-center overflow-hidden rounded-2xl bg-[#f1f4f9] sm:h-[84px]">
        {product.imageData ? (
          <img src={product.imageData} alt="" className="h-full w-full object-contain" />
        ) : (
          <ImageIcon size={24} className="text-slate-300" />
        )}
      </div>
      <div className="min-w-0">
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="rounded-lg bg-[#eef1ff] px-2 py-1 font-mono text-xs font-black text-[#4f46e5]">{product.code || 'コードなし'}</span>
              <span className="text-[11px] font-bold text-[#7067dc]">一致度 {percent(score)}</span>
              {product.hasDemoMarker && <span className="rounded bg-cyan-50 px-1.5 py-0.5 text-[10px] font-bold text-cyan-700">デモ機</span>}
            </div>
            <h4 className="mt-1.5 truncate text-sm font-bold text-[#1f2a44]">{product.name || product.itemNumber || '商品名未取得'}</h4>
          </div>
          {product.salesCount > 0 && (
            <span className="shrink-0 text-[10px] font-bold text-slate-500">実績 {product.salesCount.toLocaleString()}</span>
          )}
        </div>
        <p className="mt-1 line-clamp-2 text-[11px] leading-relaxed text-[#667085]">
          {product.catchCopy || product.specifications?.join(' / ') || product.sourceText || 'テキスト情報はまだありません。'}
        </p>
        <div className="mt-2 flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={() => onSelectSimilar(product)}
            className="rounded-full px-2.5 py-1 text-[10px] font-bold text-[#6750a4] transition hover:bg-[#f2edff]"
          >
            類似品を見る
          </button>
          {assignment && (
            <button
              type="button"
              onClick={() => onOpenSheet?.(assignment.sheetId)}
              className="flex items-center gap-1 rounded-full px-2.5 py-1 text-[10px] font-bold text-[#526070] transition hover:bg-[#eef2f7]"
            >
              <MapPin size={11} /> P.{assignment.pageNumber}を開く
            </button>
          )}
          {product.sizeType && <span className="ml-auto text-[10px] text-slate-400">{product.sizeType}</span>}
        </div>
      </div>
    </article>
  );
};

const EmptyState = ({ children }) => (
  <div className="flex min-h-48 flex-col items-center justify-center px-8 py-10 text-center text-sm text-[#7a8495]">
    <Sparkles size={24} className="mb-3 text-[#a78bfa]" />
    <span>{children}</span>
  </div>
);

const EdgeAiAssistModal = ({ isOpen, onClose, images, sheets, salesData, genres, onOpenSheet }) => {
  const [activeTab, setActiveTab] = useState('search');
  const [query, setQuery] = useState('');
  const [semanticQuery, setSemanticQuery] = useState('');
  const [semanticResults, setSemanticResults] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState(null);
  const [vectorsById, setVectorsById] = useState(null);
  const [indexFingerprint, setIndexFingerprint] = useState('');
  const [aiStatus, setAiStatus] = useState('idle');
  const [aiProgress, setAiProgress] = useState(null);
  const [aiProgressText, setAiProgressText] = useState('');
  const [aiDevice, setAiDevice] = useState('');
  const [aiError, setAiError] = useState('');
  const [oldSnapshot, setOldSnapshot] = useState(null);
  const [newSnapshot, setNewSnapshot] = useState(null);
  const [diffFilter, setDiffFilter] = useState('changed');
  const preparingRef = useRef(false);

  const products = useMemo(() => buildEdgeCatalogProducts({
    images,
    sheets,
    salesData,
    genres
  }), [genres, images, salesData, sheets]);
  const fingerprint = useMemo(() => createEdgeCatalogFingerprint(products), [products]);
  const isIndexCurrent = aiStatus === 'ready' && indexFingerprint === fingerprint && !!vectorsById;

  const lexicalResults = useMemo(() => (
    query.trim() ? rankEdgeCatalogProducts(products, query) : []
  ), [products, query]);
  const searchResults = semanticQuery === query.trim() && isIndexCurrent
    ? semanticResults
    : lexicalResults;

  const similarResults = useMemo(() => (
    selectedProduct
      ? rankSimilarEdgeCatalogProducts(products, selectedProduct, {
        vectorsById: isIndexCurrent ? vectorsById : null
      })
      : []
  ), [isIndexCurrent, products, selectedProduct, vectorsById]);

  const diffs = useMemo(() => (
    oldSnapshot && newSnapshot
      ? compareCatalogSnapshots(oldSnapshot.items, newSnapshot.items)
      : []
  ), [newSnapshot, oldSnapshot]);
  const diffSummary = useMemo(() => summarizeCatalogDiff(diffs), [diffs]);
  const visibleDiffs = useMemo(() => {
    if (diffFilter === 'all') return diffs;
    if (diffFilter === 'changed') return diffs.filter((diff) => diff.status !== 'unchanged');
    return diffs.filter((diff) => diff.status === diffFilter);
  }, [diffFilter, diffs]);

  const prepareIndex = useCallback(async () => {
    if (preparingRef.current || products.length === 0) return;
    preparingRef.current = true;
    setAiStatus('loading');
    setAiError('');
    setAiProgressText('端末内AIを起動しています…');
    try {
      const initialized = await initializeEdgeAi({ onProgress: setAiProgress });
      setAiDevice(initialized?.device || 'wasm');
      const cached = await idbHelper.getItem(INDEX_CACHE_KEY);
      if (
        cached?.fingerprint === fingerprint
        && cached?.modelVersion === EDGE_AI_MODEL_VERSION
        && cached?.vectorsById
      ) {
        setVectorsById(cached.vectorsById);
        setIndexFingerprint(fingerprint);
        setAiStatus('ready');
        setAiProgressText('保存済みの索引を読み込みました。');
        return;
      }

      const nextVectors = {};
      for (let start = 0; start < products.length; start += EMBEDDING_BATCH_SIZE) {
        const batch = products.slice(start, start + EMBEDDING_BATCH_SIZE);
        setAiProgressText(`商品を索引化しています ${Math.min(start + batch.length, products.length)} / ${products.length}`);
        const response = await embedEdgeAiTexts(
          batch.map((product) => `検索文書: ${product.searchText}`),
          { onProgress: setAiProgress }
        );
        batch.forEach((product, index) => {
          nextVectors[product.id] = response.vectors[index];
        });
      }
      await idbHelper.setItem(INDEX_CACHE_KEY, {
        modelVersion: EDGE_AI_MODEL_VERSION,
        fingerprint,
        vectorsById: nextVectors,
        savedAt: new Date().toISOString()
      });
      setVectorsById(nextVectors);
      setIndexFingerprint(fingerprint);
      setAiDevice(initialized?.device || 'wasm');
      setAiStatus('ready');
      setAiProgressText(`${products.length}商品の索引を端末に保存しました。`);
    } catch (error) {
      setAiStatus('error');
      setAiError(error instanceof Error ? error.message : '端末内AIを準備できませんでした。');
    } finally {
      preparingRef.current = false;
    }
  }, [fingerprint, products]);

  const runSearch = useCallback(async () => {
    const nextQuery = query.trim();
    if (!nextQuery) return;
    if (!isIndexCurrent) {
      setSemanticQuery('');
      setSemanticResults([]);
      return;
    }
    setAiStatus('searching');
    setAiError('');
    try {
      const response = await embedEdgeAiTexts([`検索クエリ: ${nextQuery}`]);
      setSemanticResults(rankEdgeCatalogProducts(products, nextQuery, {
        queryVector: response.vectors[0],
        vectorsById
      }));
      setSemanticQuery(nextQuery);
      setAiStatus('ready');
    } catch (error) {
      setAiStatus('error');
      setAiError(error instanceof Error ? error.message : '意味検索に失敗しました。');
    }
  }, [isIndexCurrent, products, query, vectorsById]);

  const selectSimilarProduct = useCallback((product) => {
    setSelectedProduct(product);
    setActiveTab('similar');
  }, []);

  const loadSnapshot = useCallback(async (event, setter) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    try {
      const text = await readFileAutoEncoding(file);
      setter({ ...parseCatalogSnapshotCsv(text), fileName: file.name });
    } catch (error) {
      setAiError(error instanceof Error ? error.message : 'CSVを読み込めませんでした。');
    }
  }, []);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[140] flex items-center justify-center bg-slate-950/45 p-2 backdrop-blur-sm sm:p-4" onClick={onClose}>
      <section
        role="dialog"
        aria-modal="true"
        aria-label="AIアシスト"
        className="flex h-[min(900px,96vh)] w-[min(1120px,97vw)] flex-col overflow-hidden rounded-[30px] bg-[#f7f9fc] shadow-[0_24px_80px_rgba(15,23,42,0.28)]"
        onClick={(event) => event.stopPropagation()}
      >
        <header className="flex min-h-16 items-center gap-3 bg-white/90 px-5 py-3 backdrop-blur-xl sm:px-7">
          <div className="daiwari-ai-mascot" aria-hidden="true">
            <img src="/daiwari-kun.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-front" draggable="false" />
            <img src="/daiwari-kun-left.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-left" draggable="false" />
            <img src="/daiwari-kun-closed.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-closed" draggable="false" />
          </div>
          <div className="min-w-0">
            <h2 className="text-[15px] font-bold text-[#243047]">AIアシスト <span className="ml-1 text-[10px] font-medium text-[#7568d9]">試作版</span></h2>
            <p className="hidden text-[10px] text-[#7b8597] sm:block">台割の商品情報を、この端末内だけで探して比較します</p>
          </div>
          <div className="ml-auto hidden items-center gap-1.5 text-[10px] font-medium text-emerald-700 sm:flex">
            <ShieldCheck size={13} /> 外部AI通信なし
          </div>
          <button type="button" onClick={onClose} className="rounded-full p-2 text-[#697386] transition hover:bg-[#eef1f6]" aria-label="閉じる">
            <X size={19} />
          </button>
        </header>

        <div className="flex items-center gap-1 border-b border-[#e5e9f0] bg-white/80 px-4 sm:px-7">
          <nav className="flex min-w-0 flex-1 gap-1 overflow-x-auto" aria-label="AIアシスト機能">
            {TABS.map(({ id, label }) => (
              <button
                type="button"
                key={id}
                onClick={() => setActiveTab(id)}
                className={`relative flex shrink-0 items-center gap-1.5 px-3 py-3 text-[11px] font-semibold transition ${activeTab === id ? 'text-[#5145cd]' : 'text-[#6f798b] hover:text-[#313b4d]'}`}
              >
                {id === 'search' ? <Search size={14} /> : id === 'similar' ? <Sparkles size={14} /> : <FileDiff size={14} />}
                {label}
                {activeTab === id && <span className="absolute inset-x-2 bottom-0 h-0.5 rounded-full bg-[#6557e8]" />}
              </button>
            ))}
          </nav>
          <button
            type="button"
            onClick={prepareIndex}
            disabled={!products.length || aiStatus === 'loading' || aiStatus === 'searching'}
            title={`${products.length.toLocaleString()}商品を対象。初回のみ約80MBを読み込み、索引はこのブラウザに保存します。`}
            className={`flex max-w-[220px] shrink-0 items-center gap-1.5 rounded-full px-3 py-1.5 text-[10px] font-semibold transition disabled:cursor-wait ${isIndexCurrent ? 'bg-emerald-50 text-emerald-700' : 'bg-[#eef1f6] text-[#5f697a] hover:bg-[#e5e9f0]'}`}
          >
            {aiStatus === 'loading' || aiStatus === 'searching' ? (
              <Loader2 size={12} className="shrink-0 animate-spin" />
            ) : isIndexCurrent ? (
              <CheckCircle2 size={12} className="shrink-0" />
            ) : (
              <Database size={12} className="shrink-0" />
            )}
            <span className="truncate">
              {aiStatus === 'loading'
                ? aiProgressText || formatProgress(aiProgress)
                : aiStatus === 'searching'
                  ? '検索中…'
                  : isIndexCurrent
                    ? `AI準備済み・${aiDevice.toUpperCase()}`
                    : indexFingerprint ? '索引を更新' : '意味検索を準備'}
            </span>
          </button>
        </div>

        <main className="min-h-0 flex-1 overflow-y-auto">
          <div className="mx-auto w-full max-w-[900px] px-4 py-6 sm:px-8 sm:py-8">
            {aiError && (
              <div className="mb-5 flex items-start justify-between gap-3 rounded-2xl bg-rose-50 px-4 py-3 text-xs text-rose-700">
                <span>{aiError} 軽量検索とCSV差分は引き続き使えます。</span>
                <button type="button" onClick={() => setAiError('')}><X size={14} /></button>
              </div>
            )}

            {activeTab === 'search' && (
              <div>
                {!query.trim() && (
                  <div className="pb-8 pt-7 text-center sm:pb-10 sm:pt-12">
                    <div className="daiwari-ai-mascot daiwari-ai-mascot-hero mx-auto mb-5" aria-hidden="true">
                      <img src="/daiwari-kun.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-front" draggable="false" />
                      <img src="/daiwari-kun-left.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-left" draggable="false" />
                      <img src="/daiwari-kun-closed.png" alt="" className="daiwari-ai-mascot-frame daiwari-ai-mascot-frame-closed" draggable="false" />
                    </div>
                    <h3 className="bg-gradient-to-r from-[#4e73d9] via-[#8b5cc7] to-[#d06b91] bg-clip-text text-2xl font-semibold tracking-tight text-transparent sm:text-3xl">
                      どの商品を探しますか？
                    </h3>
                    <p className="mt-2 text-xs text-[#7a8495]">介援隊コード、用途、特徴、仕様を自然な言葉で入力できます</p>
                  </div>
                )}
                <form className={`${query.trim() ? '' : 'mx-auto max-w-[760px]'} relative`} onSubmit={(event) => { event.preventDefault(); runSearch(); }}>
                  <Search size={18} className="absolute left-5 top-1/2 -translate-y-1/2 text-[#778196]" />
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    placeholder="例：折りたためる軽量な歩行器、E1423、在庫のある口腔ケア用品"
                    className="w-full rounded-[26px] border-0 bg-[#edf1f7] py-4 pl-13 pr-28 text-sm text-[#273246] shadow-none outline-none transition placeholder:text-[#8b95a6] focus:bg-white focus:ring-2 focus:ring-[#c8c2ff]"
                    autoFocus
                  />
                  <button type="submit" disabled={!query.trim() || aiStatus === 'searching'} className="absolute right-2 top-1/2 flex -translate-y-1/2 items-center gap-1 rounded-full bg-[#6254e7] px-4 py-2.5 text-[11px] font-semibold text-white transition hover:bg-[#5145cd] disabled:opacity-40">
                    {aiStatus === 'searching' ? <Loader2 size={13} className="animate-spin" /> : <ArrowRight size={13} />}
                    検索
                  </button>
                </form>
                {!query.trim() && (
                  <div className="mx-auto mt-4 flex max-w-[720px] flex-wrap justify-center gap-2">
                    {['売れ筋の車いす', '軽い歩行器', '在庫のある口腔ケア用品'].map((suggestion) => (
                      <button key={suggestion} type="button" onClick={() => setQuery(suggestion)} className="rounded-full bg-white px-3 py-1.5 text-[10px] text-[#626d80] shadow-sm transition hover:bg-[#f1edff] hover:text-[#5c4fc7]">
                        {suggestion}
                      </button>
                    ))}
                  </div>
                )}
                <div className="mt-6">
                  {!query.trim() ? (
                    <p className="text-center text-[10px] text-[#98a1b1]">
                      {isIndexCurrent ? `${products.length.toLocaleString()}商品をAI意味検索できます` : '意味検索を準備しなくても文字一致検索を利用できます'}
                    </p>
                  ) : searchResults.length === 0 ? (
                    <EmptyState>該当する商品が見つかりませんでした。別の表現でもお試しください。</EmptyState>
                  ) : (
                    <div>
                      <div className="mb-3 flex items-center justify-between px-1">
                        <p className="text-[11px] font-semibold text-[#687386]">{searchResults.length}件の候補</p>
                        <span className="text-[10px] text-[#8b95a6]">{isIndexCurrent ? 'AI意味検索' : '軽量検索'}</span>
                      </div>
                      <div className="space-y-3">
                      {searchResults.map((result) => <ResultCard key={result.product.id} result={result} onSelectSimilar={selectSimilarProduct} onOpenSheet={onOpenSheet} />)}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {activeTab === 'similar' && (
              <div>
                <h3 className="text-xl font-semibold text-[#273246]">類似品を探す</h3>
                <p className="mt-1 text-xs text-[#7a8495]">商品テキスト、仕様、コマサイズをもとに候補を並べます</p>
                {!selectedProduct ? (
                  <div className="mt-8">
                    <EmptyState>
                      <span>「商品意味検索」の結果から <b>類似品を見る</b> を選んでください。</span>
                    </EmptyState>
                  </div>
                ) : (
                  <>
                    <div className="mt-5 flex items-center gap-3 rounded-[22px] bg-[#eeebff] px-4 py-3">
                      <div className="rounded-lg bg-white px-2 py-1 font-mono text-xs font-black text-[#5d50cf]">{selectedProduct.code}</div>
                      <div className="min-w-0">
                        <p className="text-[9px] font-semibold text-[#7a6ed5]">比較元の商品</p>
                        <p className="truncate text-sm font-semibold text-[#273246]">{selectedProduct.name || selectedProduct.itemNumber || '商品名未取得'}</p>
                      </div>
                      <ArrowRight size={16} className="ml-auto text-[#8175dc]" />
                    </div>
                    <div className="mt-4 space-y-3">
                      {similarResults.map((result) => <ResultCard key={result.product.id} result={result} onSelectSimilar={selectSimilarProduct} onOpenSheet={onOpenSheet} />)}
                    </div>
                  </>
                )}
              </div>
            )}

            {activeTab === 'diff' && (
              <div>
                <h3 className="text-xl font-semibold text-[#273246]">CSVの変更を比べる</h3>
                <p className="mt-1 text-xs text-[#7a8495]">変更前と変更後の全データCSVを、介援隊コードで突合します</p>
                <div className="mt-6 grid gap-3 sm:grid-cols-2">
                  {[
                    { label: '1. 変更前CSV', snapshot: oldSnapshot, setter: setOldSnapshot },
                    { label: '2. 変更後CSV', snapshot: newSnapshot, setter: setNewSnapshot }
                  ].map(({ label, snapshot, setter }) => (
                    <label key={label} className={`flex cursor-pointer items-center gap-3 rounded-[22px] p-4 transition ${snapshot ? 'bg-emerald-50' : 'bg-white shadow-sm hover:bg-[#f2efff]'}`}>
                      {snapshot ? <CheckCircle2 size={22} className="text-emerald-600" /> : <FileUp size={22} className="text-slate-400" />}
                      <span className="min-w-0">
                        <span className="block text-xs font-semibold text-[#364154]">{label}</span>
                        <span className="block truncate text-[10px] text-[#7a8495]">{snapshot ? `${snapshot.fileName}（${snapshot.items.length.toLocaleString()}件）` : 'クリックして選択'}</span>
                      </span>
                      <input type="file" accept=".csv,text/csv" className="hidden" onChange={(event) => loadSnapshot(event, setter)} />
                    </label>
                  ))}
                </div>

                {oldSnapshot && newSnapshot ? (
                  <>
                    <div className="mt-4 flex flex-wrap gap-2">
                      {[
                        ['changed', '差分のみ', diffSummary.added + diffSummary.removed + diffSummary.modified],
                        ['modified', '変更', diffSummary.modified],
                        ['added', '追加', diffSummary.added],
                        ['removed', '削除', diffSummary.removed],
                        ['all', 'すべて', diffs.length]
                      ].map(([id, label, count]) => (
                        <button key={id} type="button" onClick={() => setDiffFilter(id)} className={`rounded-full px-3 py-1.5 text-[10px] font-semibold transition ${diffFilter === id ? 'bg-[#6254e7] text-white' : 'bg-white text-[#687386] shadow-sm hover:bg-[#efecff]'}`}>
                          {label} {count}
                        </button>
                      ))}
                    </div>
                    <div className="mt-3 space-y-2">
                      {visibleDiffs.length === 0 ? (
                        <EmptyState>選択した条件に該当する差分はありません。</EmptyState>
                      ) : visibleDiffs.map((diff) => (
                        <article key={diff.code} className="rounded-[22px] bg-white p-4 shadow-[0_1px_3px_rgba(15,23,42,0.08)]">
                          <div className="flex items-center gap-2">
                            <span className="font-mono text-sm font-black text-[#273246]">{diff.code}</span>
                            <span className={`rounded-full border px-2 py-0.5 text-[10px] font-bold ${STATUS_STYLES[diff.status]}`}>{STATUS_LABELS[diff.status]}</span>
                            <span className="truncate text-xs font-semibold text-[#566174]">{diff.after?.name || diff.before?.name || ''}</span>
                          </div>
                          {diff.changes.length > 0 && (
                            <div className="mt-3 divide-y divide-[#e7eaf0] rounded-2xl bg-[#f5f7fa] px-3">
                              {diff.changes.map((change) => (
                                <div key={change.key} className="grid gap-1 py-2 text-[11px] sm:grid-cols-[110px_1fr_18px_1fr]">
                                  <span className={`font-semibold ${change.severity === 'high' ? 'text-rose-600' : 'text-[#5f697a]'}`}>{change.label}</span>
                                  <span className="break-words text-[#7f8999] line-through">{change.before || '（空欄）'}</span>
                                  <ArrowRight size={13} className="hidden text-[#b0b7c3] sm:block" />
                                  <span className="break-words font-semibold text-[#273246]">{change.after || '（空欄）'}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </article>
                      ))}
                    </div>
                  </>
                ) : (
                  <div className="mt-4"><EmptyState>比較する2つのCSVを選択してください。</EmptyState></div>
                )}
              </div>
            )}
          </div>
        </main>
      </section>
    </div>
  );
};

export default EdgeAiAssistModal;
