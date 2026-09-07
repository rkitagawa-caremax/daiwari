import { useCallback, useMemo, useRef, useState } from 'react';
import {
  ArrowRight,
  BarChart3,
  CheckCircle2,
  CircleAlert,
  Database,
  FileDiff,
  FileUp,
  Image as ImageIcon,
  LayoutDashboard,
  Lightbulb,
  Loader2,
  MapPin,
  Search,
  ShieldCheck,
  Sparkles,
  Target,
  TrendingDown,
  TrendingUp,
  X
} from 'lucide-react';

import {
  CATALOG_DIFF_FIELD_DEFINITIONS,
  EDGE_AI_MODEL_VERSION,
  buildCatalogChangeSet,
  buildEdgeCatalogProducts,
  compareCatalogSnapshots,
  createEdgeCatalogFingerprint,
  rankEdgeCatalogProducts,
  rankSimilarEdgeCatalogProducts,
  summarizeCatalogDiff
} from '../../domain/edgeAiCatalog';
import { buildCatalogAdvisorReport } from '../../domain/catalogAdvisor';
import { embedEdgeAiTexts, initializeEdgeAi } from '../../lib/edgeAiClient';
import { readCatalogSnapshotFile } from '../../lib/catalogSnapshotFile';
import { idbHelper } from '../../idbHelper';

const INDEX_CACHE_KEY = `edgeAiCatalogIndex:${EDGE_AI_MODEL_VERSION}`;
const EMBEDDING_BATCH_SIZE = 12;

const TABS = [
  { id: 'search', label: '商品意味検索', description: '言葉で商品を探す', icon: <Search size={15} /> },
  { id: 'similar', label: '類似品提案', description: '代替候補を比較', icon: <Sparkles size={15} /> },
  { id: 'advisor', label: '台割アドバイス', description: '売上と誌面を診断', icon: <Lightbulb size={15} /> },
  { id: 'diff', label: '変更チェック', description: '更新データを照合', icon: <FileDiff size={15} /> }
];

const STATUS_LABELS = {
  added: '追加',
  removed: '削除',
  modified: '変更',
  unchanged: '変更なし'
};

const CATALOG_FIELD_LABELS = Object.fromEntries(
  CATALOG_DIFF_FIELD_DEFINITIONS.map((field) => [field.key, field.label])
);

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

const AdvisorMetric = ({ icon, label, value, note, tone = 'violet' }) => {
  const tones = {
    violet: 'bg-violet-50 text-violet-600',
    blue: 'bg-blue-50 text-blue-600',
    emerald: 'bg-emerald-50 text-emerald-600',
    amber: 'bg-amber-50 text-amber-600'
  };
  return (
    <div className="rounded-[20px] border border-white/80 bg-white p-3.5 shadow-[0_8px_24px_rgba(31,42,68,0.06)]">
      <div className="flex items-start justify-between gap-2">
        <div>
          <p className="text-[10px] font-semibold text-[#7a8495]">{label}</p>
          <p className="mt-1 font-mono text-xl font-black tracking-tight text-[#273246]">{value}</p>
        </div>
        <span className={`flex h-8 w-8 items-center justify-center rounded-xl ${tones[tone]}`}>{icon}</span>
      </div>
      {note && <p className="mt-1 text-[9px] text-[#98a1b1]">{note}</p>}
    </div>
  );
};

const AdvisorActionCard = ({ action, index }) => {
  const styles = action.priority === 'high'
    ? { badge: 'bg-rose-50 text-rose-600', rail: 'bg-rose-500', number: 'text-rose-500' }
    : action.priority === 'mid'
      ? { badge: 'bg-amber-50 text-amber-700', rail: 'bg-amber-400', number: 'text-amber-600' }
      : { badge: 'bg-blue-50 text-blue-600', rail: 'bg-blue-400', number: 'text-blue-600' };
  return (
    <article className="relative overflow-hidden rounded-[20px] border border-[#e8ebf2] bg-white p-4 pl-5 transition hover:-translate-y-0.5 hover:shadow-[0_10px_30px_rgba(31,42,68,0.08)]">
      <span className={`absolute inset-y-0 left-0 w-1 ${styles.rail}`} />
      <div className="flex items-start gap-3">
        <span className={`font-mono text-xs font-black ${styles.number}`}>{String(index + 1).padStart(2, '0')}</span>
        <div className="min-w-0 flex-1">
          <div className="flex flex-wrap items-center gap-1.5">
            <span className={`rounded-full px-2 py-0.5 text-[9px] font-bold ${styles.badge}`}>
              {action.priority === 'high' ? '優先' : action.priority === 'mid' ? '検討' : '参考'}
            </span>
            <span className="text-[9px] font-semibold text-[#8a94a6]">{action.category}</span>
            {action.metric && <span className="ml-auto font-mono text-[9px] font-bold text-[#697386]">{action.metric}</span>}
          </div>
          <h5 className="mt-2 text-xs font-bold leading-relaxed text-[#273246]">{action.title}</h5>
          <p className="mt-1 text-[10px] leading-relaxed text-[#596579]">{action.detail}</p>
          <p className="mt-2 border-t border-[#eef0f5] pt-2 text-[9px] leading-relaxed text-[#7c72c9]">{action.theory}</p>
        </div>
      </div>
    </article>
  );
};

const formatTrendChange = (row) => {
  if (!Number.isFinite(row.changeRate)) return '新規';
  const value = Math.round(Math.abs(row.changeRate) * 100);
  return `${row.changeRate >= 0 ? '+' : '-'}${value}%`;
};

const DiffResults = ({ diffs, emptyMessage = '差分はありません。' }) => (
  <div className="mt-3 space-y-2">
    {diffs.length === 0 ? (
      <EmptyState>{emptyMessage}</EmptyState>
    ) : diffs.map((diff) => (
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
);

const EdgeAiAssistModal = ({
  isOpen,
  onClose,
  images,
  sheets,
  salesData,
  genres,
  onOpenSheet,
  activeChangeSet,
  onApplyChangeSet
}) => {
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
  const [catalogSnapshot, setCatalogSnapshot] = useState(null);
  const [diffMode, setDiffMode] = useState('catalog');
  const [diffFilter, setDiffFilter] = useState('changed');
  const [advisorReport, setAdvisorReport] = useState(null);
  const preparingRef = useRef(false);

  const products = useMemo(() => buildEdgeCatalogProducts({
    images,
    sheets,
    salesData,
    genres
  }), [genres, images, salesData, sheets]);
  const fingerprint = useMemo(() => createEdgeCatalogFingerprint(products), [products]);
  const isIndexCurrent = aiStatus === 'ready' && indexFingerprint === fingerprint && !!vectorsById;

  // 台割アドバイス: 端末内で決定的に計算する (LLM もネットワークも使わない)。
  // 意味検索の索引があればカニバリ判定に埋め込み類似度を使い、無ければ語彙一致で代替する。
  const runAdvisor = useCallback(() => {
    setAdvisorReport(buildCatalogAdvisorReport({
      products,
      vectorsById: isIndexCurrent ? vectorsById : null
    }));
  }, [isIndexCurrent, products, vectorsById]);

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
  const catalogChangeSet = useMemo(() => (
    catalogSnapshot
      ? buildCatalogChangeSet({
        products,
        snapshot: catalogSnapshot,
        fileName: catalogSnapshot.fileName,
        sheetName: catalogSnapshot.sheetName
      })
      : null
  ), [catalogSnapshot, products]);

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
      setAiError('');
      setter(await readCatalogSnapshotFile(file));
    } catch (error) {
      setAiError(error instanceof Error ? error.message : '比較データを読み込めませんでした。');
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

        <div className="flex items-center gap-3 border-b border-[#e5e9f0] bg-white/80 px-3 py-2.5 sm:px-6">
          <nav className="flex min-w-0 flex-1 gap-1.5 overflow-x-auto rounded-2xl bg-[#f1f3f8] p-1.5" aria-label="AIアシスト機能" role="tablist">
            {TABS.map(({ id, label, description, icon }) => {
              const selected = activeTab === id;
              return (
                <button
                  type="button"
                  role="tab"
                  aria-selected={selected}
                  key={id}
                  onClick={() => setActiveTab(id)}
                  className={`group flex min-w-[142px] flex-1 shrink-0 items-center gap-2 rounded-xl px-2.5 py-2 text-left transition-all duration-200 ${selected ? 'bg-white text-[#5145cd] shadow-[0_5px_16px_rgba(42,50,80,0.11)] ring-1 ring-black/[0.03]' : 'text-[#697386] hover:bg-white/60 hover:text-[#313b4d]'}`}
                >
                  <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-[10px] transition ${selected ? 'bg-gradient-to-br from-[#6d5df3] to-[#8b70e8] text-white shadow-[0_4px_10px_rgba(101,87,232,0.28)]' : 'bg-white text-[#7b8496] shadow-sm group-hover:text-[#6557e8]'}`}>
                    {icon}
                  </span>
                  <span className="min-w-0">
                    <span className="block whitespace-nowrap text-[11px] font-bold">{label}</span>
                    <span className={`hidden whitespace-nowrap text-[9px] font-medium lg:block ${selected ? 'text-[#887ee0]' : 'text-[#9aa2b0]'}`}>{description}</span>
                  </span>
                </button>
              );
            })}
          </nav>
          <button
            type="button"
            onClick={prepareIndex}
            disabled={!products.length || aiStatus === 'loading' || aiStatus === 'searching'}
            title={`${products.length.toLocaleString()}商品を対象。初回のみ約80MBを読み込み、索引はこのブラウザに保存します。`}
            className={`flex max-w-[190px] shrink-0 items-center gap-1.5 rounded-xl border px-3 py-2 text-[10px] font-semibold transition disabled:cursor-wait ${isIndexCurrent ? 'border-emerald-100 bg-emerald-50 text-emerald-700' : 'border-[#e5e8ef] bg-white text-[#5f697a] shadow-sm hover:border-[#d8d2ff] hover:text-[#5c4fc7]'}`}
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
          <div className="mx-auto w-full max-w-[1020px] px-4 py-6 sm:px-7 sm:py-8">
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

            {activeTab === 'advisor' && (
              <div>
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3">
                    <span className="flex h-11 w-11 items-center justify-center rounded-2xl bg-gradient-to-br from-[#6657e8] to-[#9a75dd] text-white shadow-[0_8px_20px_rgba(101,87,232,0.24)]"><LayoutDashboard size={20} /></span>
                    <div>
                      <h3 className="text-xl font-semibold tracking-tight text-[#273246]">台割インサイト</h3>
                      <p className="mt-0.5 text-[11px] text-[#7a8495]">売上・誌面面積・品揃えを横断して、次に直す場所を見つけます</p>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={runAdvisor}
                    disabled={products.length === 0}
                    className="flex items-center gap-1.5 rounded-xl bg-[#6254e7] px-4 py-2.5 text-[11px] font-semibold text-white shadow-[0_6px_16px_rgba(98,84,231,0.22)] transition hover:-translate-y-0.5 hover:bg-[#5145cd] disabled:translate-y-0 disabled:opacity-40"
                  >
                    <BarChart3 size={14} /> {advisorReport ? '最新データで再分析' : '台割を分析する'}
                  </button>
                </div>

                {!advisorReport ? (
                  <div className="mt-6 rounded-[26px] border border-dashed border-[#d9d5f4] bg-gradient-to-br from-white to-[#f4f2ff]">
                    <EmptyState>
                      <span><b className="text-[#5145cd]">現在の台割を端末内で診断します</b><br />
                      大コマ面積や複数配置も考慮。意味検索を準備すると、重複商品の判定精度も上がります。</span>
                    </EmptyState>
                  </div>
                ) : (
                  <div className="mt-5 space-y-4">
                    <section className="relative overflow-hidden rounded-[26px] bg-gradient-to-br from-[#332d72] via-[#4c4196] to-[#6753b5] p-5 text-white shadow-[0_16px_36px_rgba(52,45,114,0.2)] sm:p-6">
                      <div className="absolute -right-12 -top-16 h-44 w-44 rounded-full bg-white/10 blur-2xl" />
                      <div className="relative flex flex-wrap items-center justify-between gap-5">
                        <div>
                          <p className="text-[9px] font-bold tracking-[0.2em] text-[#c8c1ff]">DIAGNOSIS OVERVIEW</p>
                          <h4 className="mt-2 text-xl font-semibold tracking-tight sm:text-2xl">
                            {advisorReport.summary.priorityCounts.high > 0
                              ? `優先して見直したい項目が ${advisorReport.summary.priorityCounts.high}件あります`
                              : advisorReport.summary.priorityCounts.mid > 0
                                ? `改善を検討したい項目が ${advisorReport.summary.priorityCounts.mid}件あります`
                                : '大きな偏りは見つかりませんでした'}
                          </h4>
                          <div className="mt-3 flex flex-wrap gap-2 text-[10px]">
                            <span className="rounded-full bg-white/12 px-2.5 py-1">{advisorReport.actions.length}件の提案</span>
                            <span className="rounded-full bg-white/12 px-2.5 py-1">{advisorReport.summary.genreCount}ジャンルを比較</span>
                            <span className="rounded-full bg-white/12 px-2.5 py-1">{advisorReport.summary.usedEmbeddings ? '意味ベクトル使用' : '語彙類似で判定'}</span>
                          </div>
                        </div>
                        <div className="flex min-w-[126px] items-center gap-3 rounded-2xl bg-white/10 p-3 ring-1 ring-white/15 backdrop-blur-sm">
                          <div className="relative flex h-14 w-14 items-center justify-center rounded-full bg-white/10">
                            <span className="font-mono text-lg font-black">{advisorReport.dataQuality.score}</span>
                            <span className="absolute -bottom-1 rounded-full bg-emerald-400 px-1.5 py-0.5 text-[7px] font-black text-emerald-950">DATA</span>
                          </div>
                          <div>
                            <p className="text-[9px] text-[#cec9f5]">分析信頼度</p>
                            <p className="mt-0.5 text-xs font-bold">{advisorReport.dataQuality.level === 'high' ? '高い' : advisorReport.dataQuality.level === 'medium' ? '標準' : '要データ補完'}</p>
                          </div>
                        </div>
                      </div>
                    </section>

                    <div className="grid grid-cols-2 gap-2.5 lg:grid-cols-4">
                      <AdvisorMetric icon={<LayoutDashboard size={16} />} label="掲載SKU" value={advisorReport.summary.placedCount.toLocaleString()} note={`${advisorReport.summary.placementCount.toLocaleString()}箇所に配置`} tone="violet" />
                      <AdvisorMetric icon={<Target size={16} />} label="使用コマ面積" value={advisorReport.summary.spaceUnits.toLocaleString()} note="1/16コマ換算" tone="blue" />
                      <AdvisorMetric icon={<TrendingUp size={16} />} label="期間販売数" value={advisorReport.summary.totalSales.toLocaleString()} note="現在選択中の売上データ" tone="emerald" />
                      <AdvisorMetric icon={<CheckCircle2 size={16} />} label="売上照合率" value={`${Math.round(advisorReport.summary.salesCoverage * 100)}%`} note="コード照合できた掲載SKU" tone="amber" />
                    </div>

                    <div className="grid gap-4 lg:grid-cols-[minmax(0,1.45fr)_minmax(280px,.75fr)]">
                      <section>
                        <div className="mb-2.5 flex items-center justify-between px-1">
                          <div>
                            <h4 className="text-sm font-bold text-[#273246]">次に行うこと</h4>
                            <p className="mt-0.5 text-[9px] text-[#8a94a6]">影響度の高い順に最大8件を表示</p>
                          </div>
                          {advisorReport.summary.priorityCounts.high > 0 && <span className="rounded-full bg-rose-50 px-2.5 py-1 text-[9px] font-bold text-rose-600">優先 {advisorReport.summary.priorityCounts.high}</span>}
                        </div>
                      {advisorReport.actions.length === 0 ? (
                        <div className="rounded-[20px] bg-emerald-50 p-5 text-xs text-emerald-700">現在のバランスを維持してください。</div>
                      ) : (
                        <div className="space-y-2.5">
                          {advisorReport.actions.map((action, index) => <AdvisorActionCard key={`${action.category}-${index}`} action={action} index={index} />)}
                        </div>
                      )}
                      </section>

                      <aside className="space-y-3">
                        <section className="rounded-[22px] border border-[#e8ebf2] bg-white p-4">
                          <div className="flex items-center justify-between">
                            <h4 className="text-xs font-bold text-[#273246]">ABC構成</h4>
                            <span className="text-[9px] text-[#8a94a6]">累積売上基準</span>
                          </div>
                          <div className="mt-4 flex h-2.5 overflow-hidden rounded-full bg-[#eef1f6]">
                            {['A', 'B', 'C'].map((rank) => (
                              <div key={rank} className={rank === 'A' ? 'bg-[#6557e8]' : rank === 'B' ? 'bg-[#60a5fa]' : 'bg-[#cbd5e1]'} style={{ width: `${advisorReport.summary.placedCount ? advisorReport.abc.counts[rank] / advisorReport.summary.placedCount * 100 : 0}%` }} />
                            ))}
                          </div>
                          <div className="mt-3 grid grid-cols-3 gap-2">
                            {['A', 'B', 'C'].map((rank) => (
                              <div key={rank} className="rounded-xl bg-[#f7f8fb] px-2 py-2 text-center">
                                <p className="text-[9px] font-bold text-[#8a94a6]">{rank}ランク</p>
                                <p className="mt-0.5 font-mono text-base font-black text-[#344054]">{advisorReport.abc.counts[rank]}</p>
                              </div>
                            ))}
                          </div>
                          <p className="mt-3 text-[9px] leading-relaxed text-[#7a8495]">上位10%の商品が売上の <b className="text-[#5145cd]">{Math.round(advisorReport.summary.topShare * 100)}%</b> を構成しています。</p>
                        </section>

                        <section className="rounded-[22px] border border-[#e8ebf2] bg-white p-4">
                          <h4 className="flex items-center gap-1.5 text-xs font-bold text-[#273246]"><CircleAlert size={14} className="text-[#7c6ee6]" /> 検出シグナル</h4>
                          <div className="mt-3 divide-y divide-[#eef0f5]">
                            {[
                              ['実績ゼロ', `${advisorReport.abc.zeroSalesCount}商品`],
                              ['重複候補', `${advisorReport.cannibalization.length}組`],
                              ['価格帯の不足', `${advisorReport.priceBands.rows.filter((row) => row.missing.length > 0).length}ジャンル`],
                              ['上昇 / 下降', `${advisorReport.momentum.rising.length} / ${advisorReport.momentum.falling.length}商品`]
                            ].map(([label, value]) => (
                              <div key={label} className="flex items-center justify-between py-2 text-[10px]">
                                <span className="text-[#697386]">{label}</span>
                                <span className="font-mono font-bold text-[#344054]">{value}</span>
                              </div>
                            ))}
                          </div>
                        </section>
                      </aside>
                    </div>

                    <section className="overflow-hidden rounded-[22px] border border-[#e8ebf2] bg-white">
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#edf0f5] px-4 py-3.5">
                        <div>
                          <h4 className="text-sm font-bold text-[#273246]">ジャンル別スペース効率</h4>
                          <p className="mt-0.5 text-[9px] text-[#8a94a6]">大コマを1/16単位へ換算し、複数ジャンル配置の売上を面積按分</p>
                        </div>
                        <div className="flex gap-3 text-[9px] text-[#7a8495]"><span className="flex items-center gap-1"><i className="h-2 w-2 rounded-full bg-[#a9b8f5]" />誌面面積</span><span className="flex items-center gap-1"><i className="h-2 w-2 rounded-full bg-[#34c6a3]" />売上</span></div>
                      </div>
                      <div className="divide-y divide-[#f0f2f6]">
                        {advisorReport.balance.rows.map((row) => (
                          <div key={row.genre} className="grid gap-2 px-4 py-3 text-[10px] sm:grid-cols-[130px_minmax(0,1fr)_112px] sm:items-center">
                            <div className="flex min-w-0 items-center gap-2">
                              <span className="truncate font-bold text-[#374357]">{row.genre}</span>
                              {row.status !== 'ok' && <span className={`shrink-0 rounded-full px-1.5 py-0.5 text-[8px] font-bold ${row.status === 'under' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}`}>{row.status === 'under' ? '増枠' : '縮小'}</span>}
                            </div>
                            <div className="space-y-1.5">
                              <div className="h-1.5 overflow-hidden rounded-full bg-[#eef1f6]"><div className="h-full rounded-full bg-[#a9b8f5]" style={{ width: `${Math.min(100, row.panelShare * 100)}%` }} /></div>
                              <div className="h-1.5 overflow-hidden rounded-full bg-[#eef1f6]"><div className="h-full rounded-full bg-[#34c6a3]" style={{ width: `${Math.min(100, row.salesShare * 100)}%` }} /></div>
                            </div>
                            <div className="flex items-center justify-between gap-2 font-mono text-[9px] text-[#7a8495] sm:justify-end">
                              <span>誌面{Math.round(row.panelShare * 100)}% / 売上{Math.round(row.salesShare * 100)}%</span>
                              <span className={`rounded-lg px-1.5 py-1 font-bold ${row.fairShare >= 1.5 ? 'bg-emerald-50 text-emerald-700' : row.fairShare < 0.6 ? 'bg-amber-50 text-amber-700' : 'bg-slate-50 text-slate-600'}`}>{row.fairShare.toFixed(1)}×</span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </section>

                    {(advisorReport.momentum.rising.length > 0 || advisorReport.momentum.falling.length > 0) && (
                      <div className="grid gap-3 sm:grid-cols-2">
                        {[
                          ['伸びている商品', '前3ヶ月比', advisorReport.momentum.rising, 'text-emerald-600', <TrendingUp key="up" size={14} className="text-emerald-600" />],
                          ['落ちている商品', '前3ヶ月比', advisorReport.momentum.falling, 'text-rose-500', <TrendingDown key="down" size={14} className="text-rose-500" />]
                        ].map(([title, subtitle, rows, tone, icon]) => rows.length > 0 && (
                          <section key={title} className="rounded-[22px] border border-[#e8ebf2] bg-white p-4">
                            <div className="flex items-center justify-between"><h4 className="flex items-center gap-1.5 text-xs font-bold text-[#273246]">{icon}{title}</h4><span className="text-[9px] text-[#9aa2b0]">{subtitle}</span></div>
                            <div className="mt-3 space-y-2">
                              {rows.map((row) => (
                                <div key={row.id} className="flex items-center justify-between gap-3 text-[10px]">
                                  <div className="min-w-0"><p className="truncate font-semibold text-[#4b5768]">{row.name}</p><p className="font-mono text-[8px] text-[#a0a7b4]">{Math.round(row.previousTotal)} → {Math.round(row.recentTotal)}</p></div>
                                  <span className={`shrink-0 rounded-lg bg-[#f7f8fb] px-2 py-1 font-mono font-bold ${tone}`}>{formatTrendChange(row)}</span>
                                </div>
                              ))}
                            </div>
                          </section>
                        ))}
                      </div>
                    )}

                    <section className="rounded-[20px] bg-[#eef1f6] px-4 py-3">
                      <div className="grid gap-3 sm:grid-cols-3">
                        {[
                          ['売上', advisorReport.dataQuality.salesCoverage],
                          ['月別推移', advisorReport.dataQuality.monthlyCoverage],
                          ['価格', advisorReport.dataQuality.priceCoverage]
                        ].map(([label, value]) => (
                          <div key={label}>
                            <div className="flex justify-between text-[9px] font-semibold text-[#697386]"><span>{label}データ</span><span>{Math.round(value * 100)}%</span></div>
                            <div className="mt-1 h-1 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full bg-[#7568d9]" style={{ width: `${value * 100}%` }} /></div>
                          </div>
                        ))}
                      </div>
                      <p className="mt-2.5 text-[9px] leading-relaxed text-[#8a94a6]">重複判定: {advisorReport.summary.usedEmbeddings ? '意味ベクトル＋価格差' : '語彙一致＋価格差（意味検索を準備すると精度向上）'} ／ 売上は現在選択中の期を使用 ／ 分析結果は提案であり、季節性・粗利・在庫状況と合わせて判断してください。</p>
                    </section>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'diff' && (
              <div>
                <h3 className="text-xl font-semibold text-[#273246]">商品情報の変更を確認</h3>
                <p className="mt-1 text-xs text-[#7a8495]">介援隊コードで突合し、詳細画面のコマ上に変更内容を表示します</p>

                <div className="mt-5 inline-flex rounded-full bg-[#e9edf4] p-1">
                  <button type="button" onClick={() => setDiffMode('catalog')} className={`rounded-full px-4 py-2 text-[11px] font-semibold transition ${diffMode === 'catalog' ? 'bg-white text-[#5145cd] shadow-sm' : 'text-[#667085]'}`}>台割と比較</button>
                  <button type="button" onClick={() => setDiffMode('files')} className={`rounded-full px-4 py-2 text-[11px] font-semibold transition ${diffMode === 'files' ? 'bg-white text-[#5145cd] shadow-sm' : 'text-[#667085]'}`}>2ファイル比較</button>
                </div>

                {diffMode === 'catalog' ? (
                  <>
                    <label className={`mt-5 flex cursor-pointer items-center gap-3 rounded-[22px] p-5 transition ${catalogSnapshot ? 'bg-emerald-50' : 'bg-white shadow-sm hover:bg-[#f2efff]'}`}>
                      {catalogSnapshot ? <CheckCircle2 size={24} className="text-emerald-600" /> : <FileUp size={24} className="text-slate-400" />}
                      <span className="min-w-0 flex-1">
                        <span className="block text-sm font-semibold text-[#364154]">新しいExcel／CSVを選択</span>
                        <span className="block truncate text-[10px] text-[#7a8495]">
                          {catalogSnapshot
                            ? `${catalogSnapshot.fileName}${catalogSnapshot.sheetName ? `・${catalogSnapshot.sheetName}` : ''}（${catalogSnapshot.items.length.toLocaleString()}件）`
                            : '保存済みのコマテキストを基準に比較します（.xlsx／.csv）'}
                        </span>
                      </span>
                      <input type="file" accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" className="hidden" onChange={(event) => loadSnapshot(event, setCatalogSnapshot)} />
                    </label>

                    {catalogChangeSet ? (
                      <>
                        <div className="mt-4 flex flex-wrap items-center gap-2 rounded-2xl bg-white px-4 py-3 text-[11px] text-[#667085] shadow-sm">
                          <span className="font-bold text-[#273246]">変更 {catalogChangeSet.summary.modified}</span>
                          <span>新規 {catalogChangeSet.summary.added}</span>
                          <span>照合済み {catalogSnapshot.items.length.toLocaleString()}件</span>
                          {catalogChangeSet.recognizedFields.length > 0 ? (
                            <span className="max-w-full truncate text-emerald-700">
                              認識項目 {catalogChangeSet.recognizedFields.map((key) => CATALOG_FIELD_LABELS[key] || key).join('・')}
                            </span>
                          ) : (
                            <span className="font-bold text-amber-700">比較できる項目列を認識できません</span>
                          )}
                          {catalogChangeSet.duplicateCodes > 0 && <span className="text-amber-700">重複コード {catalogChangeSet.duplicateCodes}件</span>}
                          {catalogChangeSet.unreadableRows > 0 && <span className="text-amber-700">コード不明 {catalogChangeSet.unreadableRows}行</span>}
                          <button
                            type="button"
                            onClick={() => onApplyChangeSet?.(catalogChangeSet)}
                            className="ml-auto rounded-full bg-[#6254e7] px-4 py-2 font-bold text-white transition hover:bg-[#5145cd] disabled:opacity-40"
                            disabled={catalogChangeSet.displayCount === 0}
                          >
                            {catalogChangeSet.displayCount > 0 ? `${catalogChangeSet.displayCount}件をページに表示` : 'ページ表示できる変更はありません'}
                          </button>
                        </div>
                        <DiffResults diffs={catalogChangeSet.diffs} emptyMessage="取り込んだ項目に変更はありませんでした。" />
                      </>
                    ) : activeChangeSet ? (
                      <div className="mt-4 rounded-2xl bg-indigo-50 px-4 py-3 text-[11px] text-indigo-700">
                        現在、{activeChangeSet.fileName || '取込データ'}の差分 {activeChangeSet.displayCount || Object.keys(activeChangeSet.byCode || {}).length}件をページに表示できます。
                      </div>
                    ) : (
                      <div className="mt-4"><EmptyState>新しい価格表や商品情報ファイルを選択してください。</EmptyState></div>
                    )}
                  </>
                ) : (
                  <>
                    <div className="mt-5 grid gap-3 sm:grid-cols-2">
                      {[
                        { label: '1. 変更前Excel／CSV', snapshot: oldSnapshot, setter: setOldSnapshot },
                        { label: '2. 変更後Excel／CSV', snapshot: newSnapshot, setter: setNewSnapshot }
                      ].map(({ label, snapshot, setter }) => (
                        <label key={label} className={`flex cursor-pointer items-center gap-3 rounded-[22px] p-4 transition ${snapshot ? 'bg-emerald-50' : 'bg-white shadow-sm hover:bg-[#f2efff]'}`}>
                          {snapshot ? <CheckCircle2 size={22} className="text-emerald-600" /> : <FileUp size={22} className="text-slate-400" />}
                          <span className="min-w-0">
                            <span className="block text-xs font-semibold text-[#364154]">{label}</span>
                            <span className="block truncate text-[10px] text-[#7a8495]">{snapshot ? `${snapshot.fileName}（${snapshot.items.length.toLocaleString()}件）` : 'クリックして選択'}</span>
                          </span>
                          <input type="file" accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" className="hidden" onChange={(event) => loadSnapshot(event, setter)} />
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
                        <DiffResults diffs={visibleDiffs} emptyMessage="選択した条件に該当する差分はありません。" />
                      </>
                    ) : (
                      <div className="mt-4"><EmptyState>比較する2つのExcel／CSVを選択してください。</EmptyState></div>
                    )}
                  </>
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
