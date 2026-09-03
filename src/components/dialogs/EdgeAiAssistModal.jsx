import { useCallback, useMemo, useRef, useState } from 'react';
import {
  ArrowRight,
  Bot,
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
    <article className="grid grid-cols-[76px_minmax(0,1fr)] gap-3 rounded-2xl border border-slate-200 bg-white p-3 shadow-sm transition hover:border-indigo-200 hover:shadow-md">
      <div className="flex h-[76px] items-center justify-center overflow-hidden rounded-xl bg-slate-100">
        {product.imageData ? (
          <img src={product.imageData} alt="" className="h-full w-full object-contain" />
        ) : (
          <ImageIcon size={24} className="text-slate-300" />
        )}
      </div>
      <div className="min-w-0">
        <div className="flex items-start justify-between gap-2">
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="rounded-md bg-indigo-600 px-2 py-0.5 font-mono text-xs font-black text-white">{product.code || 'コードなし'}</span>
              <span className="text-[11px] font-bold text-indigo-600">一致度 {percent(score)}</span>
              {product.hasDemoMarker && <span className="rounded bg-cyan-50 px-1.5 py-0.5 text-[10px] font-bold text-cyan-700">デモ機</span>}
            </div>
            <h4 className="mt-1 truncate text-sm font-bold text-slate-800">{product.name || product.itemNumber || '商品名未取得'}</h4>
          </div>
          {product.salesCount > 0 && (
            <span className="shrink-0 text-[10px] font-bold text-slate-500">実績 {product.salesCount.toLocaleString()}</span>
          )}
        </div>
        <p className="mt-1 line-clamp-2 text-[11px] leading-relaxed text-slate-500">
          {product.catchCopy || product.specifications?.join(' / ') || product.sourceText || 'テキスト情報はまだありません。'}
        </p>
        <div className="mt-2 flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={() => onSelectSimilar(product)}
            className="rounded-lg bg-violet-50 px-2 py-1 text-[10px] font-bold text-violet-700 hover:bg-violet-100"
          >
            類似品を見る
          </button>
          {assignment && (
            <button
              type="button"
              onClick={() => onOpenSheet?.(assignment.sheetId)}
              className="flex items-center gap-1 rounded-lg bg-slate-100 px-2 py-1 text-[10px] font-bold text-slate-600 hover:bg-slate-200"
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
  <div className="flex min-h-44 items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-slate-50/70 p-8 text-center text-sm text-slate-500">
    {children}
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
    <div className="fixed inset-0 z-[140] flex items-center justify-center bg-slate-950/55 p-3 backdrop-blur-sm" onClick={onClose}>
      <section
        role="dialog"
        aria-modal="true"
        aria-label="AIアシスト"
        className="flex h-[min(880px,94vh)] w-[min(1180px,96vw)] flex-col overflow-hidden rounded-[28px] border border-white/60 bg-slate-50 shadow-2xl"
        onClick={(event) => event.stopPropagation()}
      >
        <header className="flex items-center gap-4 border-b border-slate-200 bg-white px-6 py-4">
          <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 to-violet-600 text-white shadow-lg shadow-indigo-200">
            <Bot size={23} />
          </div>
          <div className="min-w-0">
            <h2 className="text-lg font-black text-slate-800">AIアシスト <span className="ml-1 text-xs font-bold text-indigo-500">試作版</span></h2>
            <p className="text-xs text-slate-500">商品検索・類似品・CSV変更差分を、この端末内だけで処理します。</p>
          </div>
          <div className="ml-auto hidden items-center gap-2 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-1.5 text-[11px] font-bold text-emerald-700 sm:flex">
            <ShieldCheck size={14} /> 外部AI通信なし・API課金なし
          </div>
          <button type="button" onClick={onClose} className="rounded-full p-2 text-slate-500 hover:bg-slate-100" aria-label="閉じる">
            <X size={20} />
          </button>
        </header>

        <div className="grid min-h-0 flex-1 grid-cols-1 lg:grid-cols-[230px_minmax(0,1fr)]">
          <aside className="border-b border-slate-200 bg-white p-3 lg:border-b-0 lg:border-r">
            <nav className="flex gap-2 overflow-x-auto lg:block lg:space-y-1.5">
              {TABS.map(({ id, label }) => (
                <button
                  type="button"
                  key={id}
                  onClick={() => setActiveTab(id)}
                  className={`flex shrink-0 items-center gap-2 rounded-xl px-3 py-2.5 text-xs font-bold transition lg:w-full ${activeTab === id ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-600 hover:bg-slate-100'}`}
                >
                  {id === 'search' ? <Search size={16} /> : id === 'similar' ? <Sparkles size={16} /> : <FileDiff size={16} />} {label}
                </button>
              ))}
            </nav>

            <div className="mt-3 rounded-2xl border border-indigo-100 bg-indigo-50/70 p-3">
              <div className="flex items-center justify-between gap-2">
                <span className="text-[11px] font-black text-slate-700">端末内AI</span>
                <span className={`h-2 w-2 rounded-full ${isIndexCurrent ? 'bg-emerald-500' : aiStatus === 'loading' ? 'animate-pulse bg-amber-400' : 'bg-slate-300'}`} />
              </div>
              <p className="mt-1 text-[10px] leading-relaxed text-slate-500">
                {products.length.toLocaleString()}商品を対象。初回のみ約80MBを同じアプリから読み込みます。
              </p>
              {aiStatus === 'loading' ? (
                <div className="mt-2 rounded-xl bg-white p-2">
                  <div className="flex items-center gap-2 text-[10px] font-bold text-indigo-600">
                    <Loader2 size={12} className="animate-spin" /> {aiProgressText || formatProgress(aiProgress)}
                  </div>
                  <p className="mt-1 truncate text-[9px] text-slate-400">{formatProgress(aiProgress)}</p>
                </div>
              ) : (
                <button
                  type="button"
                  onClick={prepareIndex}
                  disabled={!products.length || aiStatus === 'searching'}
                  className="mt-2 flex w-full items-center justify-center gap-1.5 rounded-xl bg-white px-2 py-2 text-[10px] font-black text-indigo-700 shadow-sm hover:bg-indigo-100 disabled:opacity-50"
                >
                  {isIndexCurrent ? <CheckCircle2 size={13} /> : <Database size={13} />}
                  {isIndexCurrent ? `AI準備済み (${aiDevice.toUpperCase()})` : indexFingerprint ? '索引を更新' : '意味検索を準備'}
                </button>
              )}
              {isIndexCurrent && aiProgressText && <p className="mt-1.5 text-[9px] text-emerald-700">{aiProgressText}</p>}
            </div>

            <p className="mt-3 px-1 text-[9px] leading-relaxed text-slate-400">
              準備前でも文字一致による軽量検索とCSV差分は利用できます。索引はFirebaseではなく、このブラウザだけに保存します。
            </p>
          </aside>

          <main className="min-h-0 overflow-y-auto p-4 sm:p-6">
            {aiError && (
              <div className="mb-4 flex items-start justify-between gap-3 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-xs text-rose-700">
                <span>{aiError} 軽量検索とCSV差分は引き続き使えます。</span>
                <button type="button" onClick={() => setAiError('')}><X size={14} /></button>
              </div>
            )}

            {activeTab === 'search' && (
              <div>
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <h3 className="text-lg font-black text-slate-800">商品意味検索</h3>
                    <p className="mt-0.5 text-xs text-slate-500">コードだけでなく、用途・特徴・仕様の言葉で探せます。</p>
                  </div>
                  <span className={`rounded-full px-2.5 py-1 text-[10px] font-bold ${isIndexCurrent ? 'bg-violet-100 text-violet-700' : 'bg-slate-200 text-slate-600'}`}>
                    {isIndexCurrent ? 'AI意味検索' : '軽量検索'}
                  </span>
                </div>
                <form className="relative mt-4" onSubmit={(event) => { event.preventDefault(); runSearch(); }}>
                  <Search size={19} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    placeholder="例：折りたためる軽量な歩行器、E1423、在庫のある口腔ケア用品"
                    className="w-full rounded-2xl border border-slate-200 bg-white py-3.5 pl-12 pr-28 text-sm shadow-sm outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-100"
                    autoFocus
                  />
                  <button type="submit" disabled={!query.trim() || aiStatus === 'searching'} className="absolute right-2 top-1/2 flex -translate-y-1/2 items-center gap-1 rounded-xl bg-indigo-600 px-4 py-2 text-xs font-bold text-white hover:bg-indigo-700 disabled:opacity-50">
                    {aiStatus === 'searching' ? <Loader2 size={13} className="animate-spin" /> : null}検索
                  </button>
                </form>
                <div className="mt-4">
                  {!query.trim() ? (
                    <EmptyState>検索したい商品のコード、用途、特徴、仕様を入力してください。</EmptyState>
                  ) : searchResults.length === 0 ? (
                    <EmptyState>該当する商品が見つかりませんでした。別の表現でもお試しください。</EmptyState>
                  ) : (
                    <div className="grid gap-3 xl:grid-cols-2">
                      {searchResults.map((result) => <ResultCard key={result.product.id} result={result} onSelectSimilar={selectSimilarProduct} onOpenSheet={onOpenSheet} />)}
                    </div>
                  )}
                </div>
              </div>
            )}

            {activeTab === 'similar' && (
              <div>
                <h3 className="text-lg font-black text-slate-800">類似品提案</h3>
                <p className="mt-0.5 text-xs text-slate-500">商品テキスト・仕様・コマサイズをもとに候補を並べます。</p>
                {!selectedProduct ? (
                  <div className="mt-4">
                    <EmptyState>
                      <span>「商品意味検索」の結果から <b>類似品を見る</b> を選んでください。</span>
                    </EmptyState>
                  </div>
                ) : (
                  <>
                    <div className="mt-4 flex items-center gap-3 rounded-2xl border border-violet-200 bg-violet-50 p-3">
                      <div className="rounded-lg bg-violet-600 px-2 py-1 font-mono text-xs font-black text-white">{selectedProduct.code}</div>
                      <div className="min-w-0">
                        <p className="text-[10px] font-bold text-violet-500">比較元</p>
                        <p className="truncate text-sm font-bold text-slate-800">{selectedProduct.name || selectedProduct.itemNumber || '商品名未取得'}</p>
                      </div>
                      <ArrowRight size={17} className="ml-auto text-violet-400" />
                    </div>
                    <div className="mt-4 grid gap-3 xl:grid-cols-2">
                      {similarResults.map((result) => <ResultCard key={result.product.id} result={result} onSelectSimilar={selectSimilarProduct} onOpenSheet={onOpenSheet} />)}
                    </div>
                  </>
                )}
              </div>
            )}

            {activeTab === 'diff' && (
              <div>
                <h3 className="text-lg font-black text-slate-800">CSV変更差分</h3>
                <p className="mt-0.5 text-xs text-slate-500">旧版と新版の全データCSVを介援隊コードで突合します。AI準備は不要です。</p>
                <div className="mt-4 grid gap-3 sm:grid-cols-2">
                  {[
                    { label: '1. 変更前CSV', snapshot: oldSnapshot, setter: setOldSnapshot },
                    { label: '2. 変更後CSV', snapshot: newSnapshot, setter: setNewSnapshot }
                  ].map(({ label, snapshot, setter }) => (
                    <label key={label} className={`flex cursor-pointer items-center gap-3 rounded-2xl border-2 border-dashed p-4 transition ${snapshot ? 'border-emerald-300 bg-emerald-50' : 'border-slate-300 bg-white hover:border-indigo-300'}`}>
                      {snapshot ? <CheckCircle2 size={22} className="text-emerald-600" /> : <FileUp size={22} className="text-slate-400" />}
                      <span className="min-w-0">
                        <span className="block text-xs font-black text-slate-700">{label}</span>
                        <span className="block truncate text-[10px] text-slate-500">{snapshot ? `${snapshot.fileName}（${snapshot.items.length.toLocaleString()}件）` : 'クリックして選択'}</span>
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
                        <button key={id} type="button" onClick={() => setDiffFilter(id)} className={`rounded-full px-3 py-1.5 text-[11px] font-bold ${diffFilter === id ? 'bg-slate-800 text-white' : 'bg-white text-slate-600 shadow-sm hover:bg-slate-100'}`}>
                          {label} {count}
                        </button>
                      ))}
                    </div>
                    <div className="mt-3 space-y-2">
                      {visibleDiffs.length === 0 ? (
                        <EmptyState>選択した条件に該当する差分はありません。</EmptyState>
                      ) : visibleDiffs.map((diff) => (
                        <article key={diff.code} className="rounded-2xl border border-slate-200 bg-white p-3 shadow-sm">
                          <div className="flex items-center gap-2">
                            <span className="font-mono text-sm font-black text-slate-800">{diff.code}</span>
                            <span className={`rounded-full border px-2 py-0.5 text-[10px] font-bold ${STATUS_STYLES[diff.status]}`}>{STATUS_LABELS[diff.status]}</span>
                            <span className="truncate text-xs font-bold text-slate-600">{diff.after?.name || diff.before?.name || ''}</span>
                          </div>
                          {diff.changes.length > 0 && (
                            <div className="mt-2 divide-y divide-slate-100 rounded-xl bg-slate-50 px-3">
                              {diff.changes.map((change) => (
                                <div key={change.key} className="grid gap-1 py-2 text-[11px] sm:grid-cols-[110px_1fr_18px_1fr]">
                                  <span className={`font-bold ${change.severity === 'high' ? 'text-rose-600' : 'text-slate-600'}`}>{change.label}</span>
                                  <span className="break-words text-slate-500 line-through">{change.before || '（空欄）'}</span>
                                  <ArrowRight size={13} className="hidden text-slate-300 sm:block" />
                                  <span className="break-words font-bold text-slate-800">{change.after || '（空欄）'}</span>
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
          </main>
        </div>
      </section>
    </div>
  );
};

export default EdgeAiAssistModal;
