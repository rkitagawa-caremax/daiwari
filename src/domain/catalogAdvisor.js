import { cosineSimilarity } from './edgeAiCatalog.js';

// --- 台割全体のアドバイス生成 (端末内・決定的な分析) ---
// LLM は使わず、掲載中の商品 (buildEdgeCatalogProducts の出力) と売上実績・埋め込みベクトルから
// ABC 分析 / フェアシェア / カニバリゼーション / 価格帯カバレッジ / 直近トレンドを計算し、
// 経営・マーケティング理論に紐づけた推奨アクションへ変換する。

export const CATALOG_ADVISOR_THRESHOLDS = Object.freeze({
  abcARatio: 0.8,             // 累積売上シェアの A ランク境界 (パレート)
  abcBRatio: 0.95,            // B ランク境界
  cannibalSimilarity: 0.92,   // 埋め込み類似度でのカニバリ判定
  cannibalJaccard: 0.55,      // ベクトル未準備時の語彙一致 (Jaccard) 判定
  weakSalesRatio: 0.2,        // 弱い方の実績が強い方の 20% 以下なら統合候補
  minStrongSales: 10,         // カニバリ判定で「強い方」に求める最低実績
  overAllocatedFairShare: 0.6, // 誌面シェア過剰 (売上シェア/誌面シェア がこの値未満)
  underAllocatedFairShare: 1.5, // 誌面シェア過少
  minGenrePanels: 6,          // バランス判定の対象にする最小コマ数
  momentumMonths: 3,          // 直近何ヶ月をトレンド判定に使うか
  momentumRiseRatio: 1.5,
  momentumFallRatio: 0.5,
  minMomentumTotal: 12,       // トレンド判定に必要な年間最低数量
  maxCannibalPairs: 12,
  maxListItems: 8
});

const T = CATALOG_ADVISOR_THRESHOLDS;

const toPrice = (value) => {
  const price = Number(value);
  return Number.isFinite(price) && price > 0 ? price : null;
};

const primaryGenre = (product) => product?.assignments?.[0]?.genre || '未設定';

const tokenize = (text) => {
  const normalized = String(text || '').normalize('NFKC').toLowerCase();
  const tokens = new Set();
  const words = normalized.match(/[a-z0-9]{2,}|[゠-ヿ぀-ゟ一-鿿]{2,}/g) || [];
  words.slice(0, 200).forEach((word) => {
    if (word.length <= 4) tokens.add(word);
    else for (let index = 0; index + 4 <= word.length; index += 2) tokens.add(word.slice(index, index + 4));
  });
  return tokens;
};

const jaccard = (left, right) => {
  if (left.size === 0 || right.size === 0) return 0;
  let intersection = 0;
  left.forEach((token) => { if (right.has(token)) intersection++; });
  return intersection / (left.size + right.size - intersection);
};

const productLabel = (product) => product.name || product.itemNumber || product.code || '(名称未取得)';

// --- ABC 分析 (パレート) ---
export const buildAbcAnalysis = (products) => {
  const ranked = [...products].sort((a, b) => (b.salesCount || 0) - (a.salesCount || 0));
  const totalSales = ranked.reduce((sum, p) => sum + (p.salesCount || 0), 0);
  let cumulative = 0;
  const counts = { A: 0, B: 0, C: 0 };
  const rankByProductId = {};
  ranked.forEach((product) => {
    const previousShare = totalSales > 0 ? cumulative / totalSales : 1;
    cumulative += product.salesCount || 0;
    // 境界を越えさせた商品も直前の累積ランクへ含める。
    // これにより、単品で売上の80%超を占める主力商品がB判定になるのを防ぐ。
    const rank = (product.salesCount || 0) === 0
      ? 'C'
      : previousShare < T.abcARatio ? 'A' : previousShare < T.abcBRatio ? 'B' : 'C';
    counts[rank]++;
    rankByProductId[product.id] = rank;
  });
  const topTenPercent = ranked.slice(0, Math.max(1, Math.ceil(ranked.length * 0.1)));
  const topShare = totalSales > 0
    ? topTenPercent.reduce((sum, p) => sum + (p.salesCount || 0), 0) / totalSales
    : 0;
  const zeroSales = ranked
    .filter((p) => (p.salesCount || 0) === 0 && p.lifecycleStatus !== '廃盤')
    .slice(0, T.maxListItems * 2);
  return { totalSales, counts, rankByProductId, topShare, zeroSales, topProducts: ranked.slice(0, T.maxListItems) };
};

// --- ジャンル別のフェアシェア (誌面シェア vs 売上シェア) ---
export const buildGenreBalance = (products) => {
  const byGenre = new Map();
  let totalPanels = 0;
  let totalSales = 0;
  products.forEach((product) => {
    const genre = primaryGenre(product);
    if (!byGenre.has(genre)) byGenre.set(genre, { genre, panels: 0, sales: 0, products: 0 });
    const entry = byGenre.get(genre);
    entry.panels += Math.max(1, product.assignments.length);
    entry.sales += product.salesCount || 0;
    entry.products++;
    totalPanels += Math.max(1, product.assignments.length);
    totalSales += product.salesCount || 0;
  });
  const rows = [...byGenre.values()].map((entry) => {
    const panelShare = totalPanels > 0 ? entry.panels / totalPanels : 0;
    const salesShare = totalSales > 0 ? entry.sales / totalSales : 0;
    const fairShare = panelShare > 0 ? salesShare / panelShare : 0;
    return {
      ...entry,
      panelShare,
      salesShare,
      fairShare,
      status: entry.panels < T.minGenrePanels || totalSales === 0
        ? 'ok'
        : fairShare < T.overAllocatedFairShare ? 'over' : fairShare > T.underAllocatedFairShare ? 'under' : 'ok'
    };
  }).sort((a, b) => b.sales - a.sales);
  return { rows, totalPanels, totalSales };
};

// --- カニバリゼーション候補 (同ジャンル内の近似商品ペア) ---
export const buildCannibalizationPairs = (products, vectorsById = null) => {
  const method = vectorsById ? 'embedding' : 'lexical';
  const groups = new Map();
  products.forEach((product) => {
    const genre = primaryGenre(product);
    if (!groups.has(genre)) groups.set(genre, []);
    groups.get(genre).push(product);
  });

  const pairs = [];
  groups.forEach((group) => {
    const candidates = group.length > 300
      ? [...group].sort((a, b) => (b.salesCount || 0) - (a.salesCount || 0)).slice(0, 300)
      : group;
    const tokens = method === 'lexical' ? candidates.map((p) => tokenize(p.searchText)) : null;
    for (let i = 0; i < candidates.length; i++) {
      for (let j = i + 1; j < candidates.length; j++) {
        const left = candidates[i];
        const right = candidates[j];
        let similarity = 0;
        if (method === 'embedding') {
          const lv = vectorsById[left.id];
          const rv = vectorsById[right.id];
          if (!lv || !rv) continue;
          similarity = cosineSimilarity(lv, rv);
          if (similarity < T.cannibalSimilarity) continue;
        } else {
          similarity = jaccard(tokens[i], tokens[j]);
          if (similarity < T.cannibalJaccard) continue;
        }
        const [strong, weak] = (left.salesCount || 0) >= (right.salesCount || 0) ? [left, right] : [right, left];
        if ((strong.salesCount || 0) < T.minStrongSales) continue;
        if ((weak.salesCount || 0) > (strong.salesCount || 0) * T.weakSalesRatio) continue;
        pairs.push({
          genre: primaryGenre(strong),
          strong: { id: strong.id, code: strong.code, name: productLabel(strong), salesCount: strong.salesCount || 0 },
          weak: { id: weak.id, code: weak.code, name: productLabel(weak), salesCount: weak.salesCount || 0 },
          similarity,
          method
        });
      }
    }
  });
  return pairs.sort((a, b) => b.similarity - a.similarity).slice(0, T.maxCannibalPairs);
};

// --- 価格帯カバレッジ (エントリー / ミドル / プレミアムの価格ラダー) ---
export const buildPriceBandCoverage = (products) => {
  const prices = products.map((p) => toPrice(p.priceIncludingTax)).filter((p) => p !== null).sort((a, b) => a - b);
  if (prices.length < 9) return { bands: null, rows: [] };
  const t1 = prices[Math.floor(prices.length / 3)];
  const t2 = prices[Math.floor((prices.length * 2) / 3)];
  const byGenre = new Map();
  products.forEach((product) => {
    const price = toPrice(product.priceIncludingTax);
    if (price === null) return;
    const genre = primaryGenre(product);
    if (!byGenre.has(genre)) byGenre.set(genre, { genre, low: 0, mid: 0, high: 0, total: 0 });
    const entry = byGenre.get(genre);
    entry[price <= t1 ? 'low' : price <= t2 ? 'mid' : 'high']++;
    entry.total++;
  });
  const rows = [...byGenre.values()]
    .filter((entry) => entry.total >= 5)
    .map((entry) => ({
      ...entry,
      missing: ['low', 'mid', 'high'].filter((band) => entry[band] === 0)
    }))
    .sort((a, b) => b.missing.length - a.missing.length || b.total - a.total);
  return { bands: { t1, t2 }, rows };
};

// --- 直近トレンド (月別実績の勢い) ---
export const buildSalesMomentum = (products) => {
  const rising = [];
  const falling = [];
  products.forEach((product) => {
    const series = Array.isArray(product.monthlySales) ? product.monthlySales : [];
    if (series.length < T.momentumMonths + 3) return;
    const total = series.reduce((sum, entry) => sum + (entry.count || 0), 0);
    if (total < T.minMomentumTotal) return;
    const recent = series.slice(-T.momentumMonths);
    const earlier = series.slice(0, series.length - T.momentumMonths);
    const recentAvg = recent.reduce((sum, entry) => sum + (entry.count || 0), 0) / recent.length;
    const earlierAvg = earlier.reduce((sum, entry) => sum + (entry.count || 0), 0) / earlier.length;
    if (earlierAvg <= 0) return;
    const ratio = recentAvg / earlierAvg;
    const row = { id: product.id, code: product.code, name: productLabel(product), ratio, total, genre: primaryGenre(product) };
    if (ratio >= T.momentumRiseRatio) rising.push(row);
    else if (ratio <= T.momentumFallRatio) falling.push(row);
  });
  rising.sort((a, b) => b.ratio - a.ratio);
  falling.sort((a, b) => a.ratio - b.ratio);
  return { rising: rising.slice(0, T.maxListItems), falling: falling.slice(0, T.maxListItems) };
};

const percent = (value) => `${Math.round(value * 100)}%`;

// 分析結果 → 優先度付きの推奨アクション (根拠となる理論を明記する)
export const buildAdvisorActions = ({ abc, balance, cannibalization, priceBands, momentum, salesCoverage }) => {
  const actions = [];

  if (abc.zeroSales.length >= 3) {
    actions.push({
      priority: 'high',
      title: `実績ゼロの掲載商品が${abc.zeroSales.length}件あります`,
      detail: `${abc.zeroSales.slice(0, 3).map((p) => productLabel(p)).join('、')} などは誌面を使いながら売上に貢献していません。差し替え・縮小 (1/16化)・カット候補として棚卸ししてください。`,
      theory: 'ABC分析: Cランク商品の整理は誌面の回転率 (スペース生産性) を直接引き上げる、小売の基本手法です。'
    });
  }

  cannibalization.forEach((pair) => {
    actions.push({
      priority: 'high',
      title: `カニバリ候補: ${pair.weak.name} は ${pair.strong.name} と役割が重複`,
      detail: `同ジャンル「${pair.genre}」内で内容がほぼ同等 (類似度${percent(pair.similarity)}) なのに実績は ${pair.weak.salesCount.toLocaleString()} vs ${pair.strong.salesCount.toLocaleString()}。弱い方を外し、空いたコマに別の価格帯・用途の商品を入れる余地があります。`,
      theory: 'カニバリゼーション回避と「選択のパラドックス」: 近似選択肢の並列は1商品あたりの購買率を下げます。差別化軸 (価格・サイズ・機能) を明確に分けるのが定石です。'
    });
  });

  balance.rows.filter((row) => row.status === 'under').forEach((row) => {
    actions.push({
      priority: 'mid',
      title: `「${row.genre}」は誌面が不足気味です`,
      detail: `売上シェア${percent(row.salesShare)}に対し誌面シェアは${percent(row.panelShare)}。需要に対して露出が追いついていないため、増コマ・大コマ化で売上の取りこぼしを防げます。`,
      theory: 'フェアシェア理論: 露出シェアを売上シェアに合わせると機会損失が最小になります (棚割の基本)。'
    });
  });
  balance.rows.filter((row) => row.status === 'over').forEach((row) => {
    actions.push({
      priority: 'mid',
      title: `「${row.genre}」は誌面過剰の可能性があります`,
      detail: `誌面シェア${percent(row.panelShare)}に対し売上シェアは${percent(row.salesShare)}。コマを絞って伸びているジャンルへ譲る検討を。`,
      theory: 'スペース生産性: 面積あたり売上の低い区画の縮小は、カタログ全体の売上効率を高めます。'
    });
  });

  priceBands.rows.filter((row) => row.missing.length > 0).slice(0, 3).forEach((row) => {
    const bandLabel = { low: 'エントリー(低価格)', mid: 'ミドル', high: 'プレミアム(高価格)' };
    actions.push({
      priority: 'mid',
      title: `「${row.genre}」に${row.missing.map((band) => bandLabel[band]).join('・')}帯がありません`,
      detail: '価格の階段が欠けると、予算の合わない顧客を逃し、比較購買 (松竹梅) も働きません。欠けた帯の商品追加を検討してください。',
      theory: '価格ラダー / 松竹梅効果: 3価格帯を揃えると中位価格の選択率が上がることが知られています。'
    });
  });

  if (momentum.rising.length > 0) {
    actions.push({
      priority: 'info',
      title: `直近で伸びている商品が${momentum.rising.length}件あります`,
      detail: `${momentum.rising.slice(0, 3).map((row) => row.name).join('、')} などは直近3ヶ月が年間平均を大きく上回っています。次号で大コマ化・前方配置・関連商品の併載を検討する価値があります。`,
      theory: 'モメンタム戦略: 伸びている商品への追加投資は、平均回帰前に需要を最大限取り込む定石です。'
    });
  }
  if (momentum.falling.length > 0) {
    actions.push({
      priority: 'info',
      title: `勢いが落ちている商品が${momentum.falling.length}件あります`,
      detail: `${momentum.falling.slice(0, 3).map((row) => row.name).join('、')} は直近3ヶ月が失速しています。季節要因か構造的な下降かを確認し、後者ならコマ縮小を。`,
      theory: 'プロダクトライフサイクル: 衰退期の商品は露出を段階的に縮小し、成長期の商品へ資源を移します。'
    });
  }

  if (abc.topShare >= 0.6) {
    actions.push({
      priority: 'info',
      title: `売上の${percent(abc.topShare)}を上位10%の商品が生んでいます`,
      detail: 'A ランク商品の欠品・廃盤リスクがカタログ全体の売上リスクです。主力の代替候補を1つずつ用意しつつ、A 商品は目立つ位置と十分なコマサイズを維持してください。',
      theory: 'パレートの法則 (80:20): 集中は効率的ですが、上位依存はリスク管理とセットで運用します。'
    });
  }
  if (salesCoverage < 0.5) {
    actions.push({
      priority: 'info',
      title: '売上データと照合できた商品が半数未満です',
      detail: '介援隊コードの記載漏れや売上CSVの期間ずれがあると、この分析の精度が下がります。コード整備と最新CSVの取り込みを先に行うと判断材料が揃います。',
      theory: 'データドリブン経営の前提は計測です。まず照合率を上げることが最も費用対効果の高い一手です。'
    });
  }

  const order = { high: 0, mid: 1, info: 2 };
  return actions.sort((a, b) => order[a.priority] - order[b.priority]);
};

// エントリポイント: 掲載中の商品だけを対象に全セクションを計算する
export const buildCatalogAdvisorReport = ({ products = [], vectorsById = null } = {}) => {
  const placed = products.filter((product) => (product.assignments || []).length > 0);
  const withSales = placed.filter((product) => (product.salesCount || 0) > 0);
  const salesCoverage = placed.length > 0 ? withSales.length / placed.length : 0;

  const abc = buildAbcAnalysis(placed);
  const balance = buildGenreBalance(placed);
  const cannibalization = buildCannibalizationPairs(placed, vectorsById);
  const priceBands = buildPriceBandCoverage(placed);
  const momentum = buildSalesMomentum(placed);
  const actions = buildAdvisorActions({ abc, balance, cannibalization, priceBands, momentum, salesCoverage });

  return {
    generatedAt: new Date().toISOString(),
    summary: {
      placedCount: placed.length,
      totalSales: abc.totalSales,
      salesCoverage,
      topShare: abc.topShare,
      genreCount: balance.rows.length,
      usedEmbeddings: !!vectorsById
    },
    abc,
    balance,
    cannibalization,
    priceBands,
    momentum,
    actions
  };
};
