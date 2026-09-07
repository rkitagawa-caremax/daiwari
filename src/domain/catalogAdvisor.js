import { cosineSimilarity } from './edgeAiCatalog.js';

// --- 台割全体のアドバイス生成 (端末内・決定的な分析) ---
// LLM は使わず、掲載中の商品 (buildEdgeCatalogProducts の出力) と販売実績・誌面テキスト・埋め込みベクトルから
// ABC 分析 / フェアシェア / カニバリゼーション / 価格帯カバレッジ / 直近トレンドを計算し、
// 経営・マーケティング理論に紐づけた推奨アクションへ変換する。

export const CATALOG_ADVISOR_THRESHOLDS = Object.freeze({
  abcARatio: 0.8,             // 累積販売数量シェアの A ランク境界 (パレート)
  abcBRatio: 0.95,            // B ランク境界
  cannibalSimilarity: 0.92,   // 埋め込み類似度でのカニバリ判定
  cannibalJaccard: 0.55,      // ベクトル未準備時の語彙一致 (Jaccard) 判定
  weakSalesRatio: 0.2,        // 弱い方の実績が強い方の 20% 以下なら統合候補
  minStrongSales: 10,         // カニバリ判定で「強い方」に求める最低実績
  maxCannibalPriceRatio: 1.75, // 価格差が大きい商品は役割が異なるため除外
  overAllocatedFairShare: 0.6, // 誌面シェア過剰 (販売数量シェア/誌面シェア がこの値未満)
  underAllocatedFairShare: 1.5, // 誌面シェア過少
  minGenrePanels: 6,          // バランス判定の対象にする最小誌面ユニット数
  priceLowRatio: 0.75,        // ジャンル中央値に対するエントリー帯
  priceHighRatio: 1.5,        // ジャンル中央値に対するプレミアム帯
  momentumMonths: 3,          // 直近何ヶ月をトレンド判定に使うか
  momentumRiseRatio: 1.5,
  momentumFallRatio: 0.5,
  minMomentumTotal: 12,       // トレンド判定に必要な年間最低数量
  minMomentumWindowSales: 6,  // 直近または比較期間に求める最低数量
  maxCannibalCandidatesPerGenre: 240,
  maxCannibalPairs: 12,
  maxListItems: 8,
  maxActions: 8
});

const T = CATALOG_ADVISOR_THRESHOLDS;

const toPrice = (value) => {
  const price = Number(value);
  return Number.isFinite(price) && price > 0 ? price : null;
};

const primaryGenre = (product) => product?.assignments?.[0]?.genre || '未設定';

const assignmentSpace = (assignment) => (
  Math.max(1, Number(assignment?.rowSpan) || 1) * Math.max(1, Number(assignment?.colSpan) || 1)
);

const assignmentKey = (assignment) => `${assignment?.sheetId || ''}:${assignment?.panelIndex ?? ''}`;

// 同一コードの画像が複数登録されていても、SKU実績を重複加算しない。
export const consolidateAdvisorProducts = (products = []) => {
  const bySku = new Map();
  products.forEach((product) => {
    if (!product || !(product.assignments || []).length) return;
    const key = product.code ? `code:${product.code}` : `id:${product.id}`;
    const current = bySku.get(key);
    if (!current) {
      bySku.set(key, { ...product, assignments: [...product.assignments] });
      return;
    }
    const assignments = [...current.assignments, ...product.assignments];
    const seen = new Set();
    current.assignments = assignments.filter((assignment) => {
      const value = assignmentKey(assignment);
      if (seen.has(value)) return false;
      seen.add(value);
      return true;
    });
    if (current.salesMatched != null || product.salesMatched != null) {
      current.salesMatched = current.salesMatched === true || product.salesMatched === true;
    }
    if ((product.salesCount || 0) > (current.salesCount || 0)) current.salesCount = product.salesCount;
    if ((product.salesAmount || 0) > (current.salesAmount || 0)) current.salesAmount = product.salesAmount;
    if ((product.grossProfitAmount || 0) > (current.grossProfitAmount || 0)) current.grossProfitAmount = product.grossProfitAmount;
    ['quantityMatched', 'salesAmountMatched', 'grossProfitMatched'].forEach((keyName) => {
      if (current[keyName] != null || product[keyName] != null) {
        current[keyName] = current[keyName] === true || product[keyName] === true;
      }
    });
    current.catalogTextCompleteness = Math.max(current.catalogTextCompleteness || 0, product.catalogTextCompleteness || 0);
    current.catalogText = [...new Set([current.catalogText, product.catalogText].filter(Boolean))].join(' ');
    if ((product.monthlySales || []).length > (current.monthlySales || []).length) current.monthlySales = product.monthlySales;
    if (!current.priceIncludingTax && product.priceIncludingTax) current.priceIncludingTax = product.priceIncludingTax;
    if (!current.name && product.name) current.name = product.name;
    current.searchText = [...new Set([current.searchText, product.searchText].filter(Boolean))].join(' ');
  });
  return [...bySku.values()];
};

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

const PERFORMANCE_METRICS = Object.freeze([
  Object.freeze({ id: 'quantity', valueKey: 'salesCount', matchedKey: 'quantityMatched', weight: 0.3 }),
  Object.freeze({ id: 'salesAmount', valueKey: 'salesAmount', matchedKey: 'salesAmountMatched', weight: 0.3 }),
  Object.freeze({ id: 'grossProfitAmount', valueKey: 'grossProfitAmount', matchedKey: 'grossProfitMatched', weight: 0.4 })
]);

const productHasMetric = (product, metric) => (
  product?.[metric.matchedKey] === true
  || (product?.[metric.matchedKey] == null
    && Number.isFinite(Number(product?.[metric.valueKey]))
    && Number(product?.[metric.valueKey]) !== 0)
);

const resolveAvailablePerformanceMetrics = (products) => PERFORMANCE_METRICS.filter((metric) => (
  products.some((product) => productHasMetric(product, metric))
));

const performanceMetricWeights = (metrics) => {
  const totalWeight = metrics.reduce((sum, metric) => sum + metric.weight, 0) || 1;
  return Object.fromEntries(metrics.map((metric) => [metric.id, metric.weight / totalWeight]));
};

const productPerformanceBasis = (product) => {
  const candidates = [
    { metric: 'grossProfitAmount', label: '粗利額', value: Number(product?.grossProfitAmount) || 0, matched: productHasMetric(product, PERFORMANCE_METRICS[2]) },
    { metric: 'salesAmount', label: '売上額', value: Number(product?.salesAmount) || 0, matched: productHasMetric(product, PERFORMANCE_METRICS[1]) },
    { metric: 'quantity', label: '販売数量', value: Number(product?.salesCount) || 0, matched: productHasMetric(product, PERFORMANCE_METRICS[0]) }
  ];
  return candidates.find((candidate) => candidate.matched) || candidates[2];
};

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
    // これにより、単品で販売数量の80%超を占める主力商品がB判定になるのを防ぐ。
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
  const allZeroSales = ranked.filter((p) => p.salesMatched !== false && (p.salesCount || 0) === 0 && p.lifecycleStatus !== '廃盤');
  return {
    totalSales,
    counts,
    rankByProductId,
    topShare,
    zeroSalesCount: allZeroSales.length,
    zeroSales: allZeroSales.slice(0, T.maxListItems * 2),
    topProducts: ranked.slice(0, T.maxListItems)
  };
};

// --- 商品コマ別の販売数量と誌面効率 ---
export const buildPanelQuantityAnalysis = (products) => {
  const rows = products
    .filter((product) => product.salesMatched !== false)
    .map((product) => {
      const spaceUnits = (product.assignments || []).reduce((sum, assignment) => sum + assignmentSpace(assignment), 0) || 1;
      const quantity = product.salesCount || 0;
      return {
        id: product.id,
        code: product.code,
        name: productLabel(product),
        quantity,
        spaceUnits,
        quantityPerSpace: quantity / spaceUnits,
        genre: primaryGenre(product),
        assignment: product.assignments?.[0] || null
      };
    });
  const totalQuantity = rows.reduce((sum, row) => sum + row.quantity, 0);
  const totalSpaceUnits = rows.reduce((sum, row) => sum + row.spaceUnits, 0);
  const averagePerSpace = totalSpaceUnits > 0 ? totalQuantity / totalSpaceUnits : 0;
  const averagePerProduct = rows.length > 0 ? totalQuantity / rows.length : 0;
  const highest = [...rows]
    .sort((a, b) => b.quantity - a.quantity || b.quantityPerSpace - a.quantityPerSpace)
    .slice(0, T.maxListItems);
  const lowest = [...rows]
    .sort((a, b) => a.quantity - b.quantity || b.spaceUnits - a.spaceUnits)
    .slice(0, T.maxListItems);
  const expansionCandidates = rows
    .filter((row) => row.quantity >= averagePerProduct && row.quantityPerSpace >= averagePerSpace * 1.5 && row.spaceUnits <= 4)
    .sort((a, b) => b.quantityPerSpace - a.quantityPerSpace)
    .slice(0, 3);
  const reductionCandidates = rows
    .filter((row) => row.spaceUnits >= 2 && row.quantityPerSpace <= averagePerSpace * 0.5)
    .sort((a, b) => a.quantityPerSpace - b.quantityPerSpace)
    .slice(0, 3);
  return { rows, highest, lowest, expansionCandidates, reductionCandidates, averagePerSpace };
};

// --- 商品コマ別の総合実績 (数量・売上額・粗利額を、利用可能な指標だけで統合) ---
export const buildPanelPerformanceAnalysis = (products) => {
  const candidateMetrics = resolveAvailablePerformanceMetrics(products);
  const metricTotals = Object.fromEntries(candidateMetrics.map((metric) => [
    metric.id,
    products.reduce((sum, product) => sum + (productHasMetric(product, metric) ? Math.max(0, Number(product[metric.valueKey]) || 0) : 0), 0)
  ]));
  const availableMetrics = candidateMetrics.filter((metric) => metricTotals[metric.id] > 0);
  const weights = performanceMetricWeights(availableMetrics);
  const rows = products
    .filter((product) => availableMetrics.some((metric) => productHasMetric(product, metric)))
    .map((product) => {
      const spaceUnits = (product.assignments || []).reduce((sum, assignment) => sum + assignmentSpace(assignment), 0) || 1;
      const metricShares = {};
      let weightedShare = 0;
      availableMetrics.forEach((metric) => {
        if (!productHasMetric(product, metric)) return;
        metricShares[metric.id] = Math.max(0, Number(product[metric.valueKey]) || 0) / metricTotals[metric.id];
        weightedShare += metricShares[metric.id] * weights[metric.id];
      });
      return {
        id: product.id,
        code: product.code,
        name: productLabel(product),
        quantity: Number(product.salesCount) || 0,
        salesAmount: Number(product.salesAmount) || 0,
        grossProfitAmount: Number(product.grossProfitAmount) || 0,
        grossMargin: Number(product.salesAmount) !== 0 ? (Number(product.grossProfitAmount) || 0) / Number(product.salesAmount) : null,
        spaceUnits,
        metricShares,
        rawPerformanceShare: weightedShare,
        performanceShare: weightedShare,
        genre: primaryGenre(product),
        assignment: product.assignments?.[0] || null
      };
    });
  // 商品ごとに欠けている指標があっても、個別に重みを掛け直すと全体シェアが100%を超える。
  // 全商品で最後に正規化し、フェアシェアの基準を常に一貫させる。
  const totalPerformanceShare = rows.reduce((sum, row) => sum + row.rawPerformanceShare, 0);
  rows.forEach((row) => {
    row.performanceShare = totalPerformanceShare > 0 ? row.rawPerformanceShare / totalPerformanceShare : 0;
    delete row.rawPerformanceShare;
  });
  const totalSpaceUnits = rows.reduce((sum, row) => sum + row.spaceUnits, 0);
  rows.forEach((row) => {
    row.panelShare = totalSpaceUnits > 0 ? row.spaceUnits / totalSpaceUnits : 0;
    row.fairShare = row.panelShare > 0 ? row.performanceShare / row.panelShare : 0;
  });
  const highest = [...rows].sort((a, b) => b.performanceShare - a.performanceShare || b.fairShare - a.fairShare).slice(0, T.maxListItems);
  const lowest = [...rows].sort((a, b) => a.performanceShare - b.performanceShare || a.fairShare - b.fairShare).slice(0, T.maxListItems);
  const expansionCandidates = rows.filter((row) => row.fairShare >= T.underAllocatedFairShare && row.spaceUnits <= 4)
    .sort((a, b) => b.fairShare - a.fairShare).slice(0, 3);
  const reductionCandidates = rows.filter((row) => row.spaceUnits >= 2 && row.fairShare <= T.overAllocatedFairShare)
    .sort((a, b) => a.fairShare - b.fairShare).slice(0, 3);
  return { rows, highest, lowest, expansionCandidates, reductionCandidates, availableMetrics: availableMetrics.map((metric) => metric.id), weights, metricTotals };
};

// --- ジャンル別のフェアシェア (誌面シェア vs 数量・売上額・粗利額の総合実績シェア) ---
export const buildGenreBalance = (products) => {
  const byGenre = new Map();
  let totalSpaceUnits = 0;
  let totalSales = 0;
  let totalSalesAmount = 0;
  let totalGrossProfitAmount = 0;
  const candidateMetrics = resolveAvailablePerformanceMetrics(products);
  products.forEach((product) => {
    const assignments = product.assignments?.length ? product.assignments : [{ genre: '未設定' }];
    const productSpace = assignments.reduce((sum, assignment) => sum + assignmentSpace(assignment), 0);
    const quantity = productHasMetric(product, PERFORMANCE_METRICS[0]) ? Number(product.salesCount) || 0 : 0;
    const salesAmount = productHasMetric(product, PERFORMANCE_METRICS[1]) ? Number(product.salesAmount) || 0 : 0;
    const grossProfitAmount = productHasMetric(product, PERFORMANCE_METRICS[2]) ? Number(product.grossProfitAmount) || 0 : 0;
    assignments.forEach((assignment) => {
      const genre = assignment.genre || '未設定';
      if (!byGenre.has(genre)) byGenre.set(genre, { genre, spaceUnits: 0, placements: 0, sales: 0, salesAmount: 0, grossProfitAmount: 0, productIds: new Set() });
      const entry = byGenre.get(genre);
      const spaceUnits = assignmentSpace(assignment);
      entry.spaceUnits += spaceUnits;
      entry.placements++;
      entry.sales += quantity * (spaceUnits / productSpace);
      entry.salesAmount += salesAmount * (spaceUnits / productSpace);
      entry.grossProfitAmount += grossProfitAmount * (spaceUnits / productSpace);
      entry.productIds.add(product.id);
      totalSpaceUnits += spaceUnits;
    });
    totalSales += quantity;
    totalSalesAmount += salesAmount;
    totalGrossProfitAmount += grossProfitAmount;
  });
  const metricTotals = {
    quantity: totalSales,
    salesAmount: totalSalesAmount,
    grossProfitAmount: totalGrossProfitAmount
  };
  const availableMetrics = candidateMetrics.filter((metric) => metricTotals[metric.id] > 0);
  const weights = performanceMetricWeights(availableMetrics);
  const rows = [...byGenre.values()].map((entry) => {
    const panelShare = totalSpaceUnits > 0 ? entry.spaceUnits / totalSpaceUnits : 0;
    const salesShare = totalSales > 0 ? entry.sales / totalSales : 0;
    const salesAmountShare = totalSalesAmount > 0 ? entry.salesAmount / totalSalesAmount : 0;
    const grossProfitShare = totalGrossProfitAmount > 0 ? entry.grossProfitAmount / totalGrossProfitAmount : 0;
    const metricShares = { quantity: salesShare, salesAmount: salesAmountShare, grossProfitAmount: grossProfitShare };
    const performanceShare = availableMetrics.reduce((sum, metric) => sum + metricShares[metric.id] * weights[metric.id], 0);
    const fairShare = panelShare > 0 ? performanceShare / panelShare : 0;
    return {
      genre: entry.genre,
      panels: entry.spaceUnits,
      spaceUnits: entry.spaceUnits,
      placements: entry.placements,
      products: entry.productIds.size,
      panelShare,
      salesShare,
      quantityShare: salesShare,
      salesAmountShare,
      grossProfitShare,
      performanceShare,
      fairShare,
      sales: entry.sales,
      salesPerSpace: entry.spaceUnits > 0 ? entry.sales / entry.spaceUnits : 0,
      status: entry.spaceUnits < T.minGenrePanels || availableMetrics.length === 0
        ? 'ok'
        : fairShare < T.overAllocatedFairShare ? 'over' : fairShare > T.underAllocatedFairShare ? 'under' : 'ok'
    };
  }).sort((a, b) => b.performanceShare - a.performanceShare);
  return { rows, totalPanels: totalSpaceUnits, totalSpaceUnits, totalSales, totalSalesAmount, totalGrossProfitAmount, availableMetrics: availableMetrics.map((metric) => metric.id), weights };
};

export const buildProfitabilityAnalysis = (products) => {
  const rows = products
    .filter((product) => productHasMetric(product, PERFORMANCE_METRICS[1]) && productHasMetric(product, PERFORMANCE_METRICS[2]) && Number(product.salesAmount) > 0)
    .map((product) => ({
      id: product.id,
      code: product.code,
      name: productLabel(product),
      salesAmount: Number(product.salesAmount) || 0,
      grossProfitAmount: Number(product.grossProfitAmount) || 0,
      grossMargin: (Number(product.grossProfitAmount) || 0) / Number(product.salesAmount),
      assignment: product.assignments?.[0] || null
    }));
  const totalSalesAmount = rows.reduce((sum, row) => sum + row.salesAmount, 0);
  const totalGrossProfitAmount = rows.reduce((sum, row) => sum + row.grossProfitAmount, 0);
  const grossMargin = totalSalesAmount > 0 ? totalGrossProfitAmount / totalSalesAmount : null;
  const lowMarginThreshold = grossMargin == null ? null : Math.max(0.05, grossMargin * 0.6);
  const lowMargin = lowMarginThreshold == null ? [] : rows
    .filter((row) => row.grossMargin < lowMarginThreshold && row.salesAmount >= (totalSalesAmount / Math.max(1, rows.length)))
    .sort((a, b) => a.grossMargin - b.grossMargin || b.salesAmount - a.salesAmount)
    .slice(0, T.maxListItems);
  const highGrossProfit = [...rows].sort((a, b) => b.grossProfitAmount - a.grossProfitAmount).slice(0, T.maxListItems);
  return { rows, totalSalesAmount, totalGrossProfitAmount, grossMargin, lowMarginThreshold, lowMargin, highGrossProfit };
};

// 売上・粗利を掲載場所へ帰属させる。複数ページに同じSKUがある場合は重複加算せず、掲載箇所へ均等配分する。
export const buildContributionRankings = (products = []) => {
  const pages = new Map();
  const panels = new Map();
  let totalSalesAmount = 0;
  let totalGrossProfitAmount = 0;

  products.forEach((product) => {
    const assignments = Array.isArray(product?.assignments) ? product.assignments : [];
    if (assignments.length === 0) return;
    const salesAmount = productHasMetric(product, PERFORMANCE_METRICS[1])
      ? Math.max(0, Number(product.salesAmount) || 0)
      : 0;
    const grossProfitAmount = productHasMetric(product, PERFORMANCE_METRICS[2])
      ? Math.max(0, Number(product.grossProfitAmount) || 0)
      : 0;
    if (salesAmount <= 0 && grossProfitAmount <= 0) return;

    totalSalesAmount += salesAmount;
    totalGrossProfitAmount += grossProfitAmount;
    const allocation = 1 / assignments.length;
    assignments.forEach((assignment) => {
      const pageKey = String(assignment.sheetId || `page:${assignment.pageNumber || 'unknown'}`);
      if (!pages.has(pageKey)) {
        pages.set(pageKey, {
          id: pageKey,
          sheetId: assignment.sheetId || null,
          pageNumber: Number(assignment.pageNumber) || null,
          salesAmount: 0,
          grossProfitAmount: 0,
          productIds: new Set(),
          genres: new Set()
        });
      }
      const page = pages.get(pageKey);
      page.salesAmount += salesAmount * allocation;
      page.grossProfitAmount += grossProfitAmount * allocation;
      page.productIds.add(product.id);
      if (assignment.genre) page.genres.add(assignment.genre);

      const panelKey = assignmentKey(assignment);
      if (!panels.has(panelKey)) {
        panels.set(panelKey, {
          id: panelKey,
          sheetId: assignment.sheetId || null,
          pageNumber: Number(assignment.pageNumber) || null,
          panelIndex: Number.isInteger(Number(assignment.panelIndex)) ? Number(assignment.panelIndex) : null,
          genre: assignment.genre || '未設定',
          salesAmount: 0,
          grossProfitAmount: 0,
          productIds: new Set(),
          productNames: []
        });
      }
      const panel = panels.get(panelKey);
      panel.salesAmount += salesAmount * allocation;
      panel.grossProfitAmount += grossProfitAmount * allocation;
      if (!panel.productIds.has(product.id)) panel.productNames.push(productLabel(product));
      panel.productIds.add(product.id);
    });
  });

  const pageRows = [...pages.values()].map(({ productIds, genres, ...page }) => ({
    ...page,
    productCount: productIds.size,
    genres: [...genres]
  }));
  const panelRows = [...panels.values()].map(({ productIds, productNames, ...panel }) => ({
    ...panel,
    productCount: productIds.size,
    name: productNames[0] || '商品名未取得'
  }));
  const rank = (rows, key) => [...rows]
    .filter((row) => row[key] > 0)
    .sort((left, right) => right[key] - left[key]
      || (left.pageNumber || Number.MAX_SAFE_INTEGER) - (right.pageNumber || Number.MAX_SAFE_INTEGER)
      || (left.panelIndex ?? Number.MAX_SAFE_INTEGER) - (right.panelIndex ?? Number.MAX_SAFE_INTEGER))
    .slice(0, T.maxListItems);

  return {
    totals: { salesAmount: totalSalesAmount, grossProfitAmount: totalGrossProfitAmount },
    topSalesPages: rank(pageRows, 'salesAmount'),
    topGrossProfitPages: rank(pageRows, 'grossProfitAmount'),
    topSalesPanels: rank(panelRows, 'salesAmount'),
    topGrossProfitPanels: rank(panelRows, 'grossProfitAmount')
  };
};

export const buildCatalogTextAnalysis = (products, panelPerformance = null) => {
  const performanceById = new Map((panelPerformance?.rows || []).map((row) => [row.id, row.performanceShare]));
  const averagePerformanceShare = panelPerformance?.rows?.length ? 1 / panelPerformance.rows.length : 0;
  const rows = products.map((product) => {
    const fallbackSignals = [product.name, product.itemNumber, product.catchCopy, product.sourceText, product.searchText].filter((value) => String(value || '').trim()).length;
    const completeness = Number.isFinite(Number(product.catalogTextCompleteness))
      ? Math.max(0, Math.min(1, Number(product.catalogTextCompleteness)))
      : Math.min(1, fallbackSignals / 5);
    return {
      id: product.id,
      code: product.code,
      name: productLabel(product),
      completeness,
      performanceShare: performanceById.get(product.id) || 0,
      assignment: product.assignments?.[0] || null
    };
  });
  const averageCompleteness = rows.length > 0 ? rows.reduce((sum, row) => sum + row.completeness, 0) / rows.length : 0;
  const coverage = rows.length > 0 ? rows.filter((row) => row.completeness >= 0.5).length / rows.length : 0;
  const weakHighValue = rows.filter((row) => row.completeness < 0.5 && row.performanceShare >= averagePerformanceShare)
    .sort((a, b) => b.performanceShare - a.performanceShare).slice(0, T.maxListItems);
  return { rows, coverage, averageCompleteness, weakHighValue };
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
  const keepStrongestPair = (pair) => {
    if (pairs.length < T.maxCannibalPairs) {
      pairs.push(pair);
      return;
    }
    let weakestIndex = 0;
    for (let index = 1; index < pairs.length; index += 1) {
      if (pairs[index].similarity < pairs[weakestIndex].similarity) weakestIndex = index;
    }
    if (pair.similarity > pairs[weakestIndex].similarity) pairs[weakestIndex] = pair;
  };
  groups.forEach((group) => {
    const candidates = group.length > T.maxCannibalCandidatesPerGenre
      ? [...group]
        .sort((a, b) => productPerformanceBasis(b).value - productPerformanceBasis(a).value)
        .slice(0, T.maxCannibalCandidatesPerGenre)
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
        const leftBasis = productPerformanceBasis(left);
        const rightBasis = productPerformanceBasis(right);
        const comparableMetric = leftBasis.metric === rightBasis.metric ? leftBasis.metric : 'quantity';
        const metricConfig = PERFORMANCE_METRICS.find((metric) => metric.id === comparableMetric) || PERFORMANCE_METRICS[0];
        const leftValue = Number(left[metricConfig.valueKey]) || 0;
        const rightValue = Number(right[metricConfig.valueKey]) || 0;
        const [strong, weak] = leftValue >= rightValue ? [left, right] : [right, left];
        const strongValue = Math.max(leftValue, rightValue);
        const weakValue = Math.min(leftValue, rightValue);
        const strongPrice = toPrice(strong.priceIncludingTax);
        const weakPrice = toPrice(weak.priceIncludingTax);
        if (strongPrice && weakPrice && Math.max(strongPrice, weakPrice) / Math.min(strongPrice, weakPrice) > T.maxCannibalPriceRatio) continue;
        if (strongValue < (comparableMetric === 'quantity' ? T.minStrongSales : 1)) continue;
        if (weakValue > strongValue * T.weakSalesRatio) continue;
        keepStrongestPair({
          genre: primaryGenre(strong),
          basis: comparableMetric,
          basisLabel: metricConfig.id === 'grossProfitAmount' ? '粗利額' : metricConfig.id === 'salesAmount' ? '売上額' : '販売数量',
          strong: { id: strong.id, code: strong.code, name: productLabel(strong), salesCount: strong.salesCount || 0, salesAmount: strong.salesAmount || 0, grossProfitAmount: strong.grossProfitAmount || 0, performanceValue: strongValue },
          weak: { id: weak.id, code: weak.code, name: productLabel(weak), salesCount: weak.salesCount || 0, salesAmount: weak.salesAmount || 0, grossProfitAmount: weak.grossProfitAmount || 0, performanceValue: weakValue },
          similarity,
          method
        });
      }
    }
  });
  return pairs.sort((a, b) => b.similarity - a.similarity);
};

// --- 価格帯カバレッジ (エントリー / ミドル / プレミアムの価格ラダー) ---
export const buildPriceBandCoverage = (products) => {
  const byGenre = new Map();
  products.forEach((product) => {
    const price = toPrice(product.priceIncludingTax);
    if (price === null) return;
    const genre = primaryGenre(product);
    if (!byGenre.has(genre)) byGenre.set(genre, []);
    byGenre.get(genre).push(price);
  });
  const rows = [...byGenre.entries()]
    .filter(([, prices]) => prices.length >= 5)
    .map(([genre, values]) => {
      const prices = [...values].sort((a, b) => a - b);
      const median = prices[Math.floor(prices.length / 2)];
      const lowThreshold = median * T.priceLowRatio;
      const highThreshold = median * T.priceHighRatio;
      const entry = { genre, low: 0, mid: 0, high: 0, total: prices.length, median, lowThreshold, highThreshold };
      prices.forEach((price) => {
        entry[price <= lowThreshold ? 'low' : price >= highThreshold ? 'high' : 'mid']++;
      });
      return {
        ...entry,
        missing: ['low', 'mid', 'high'].filter((band) => entry[band] === 0)
      };
    })
    .sort((a, b) => b.missing.length - a.missing.length || b.total - a.total);
  return {
    bands: { method: 'genre-relative', lowRatio: T.priceLowRatio, highRatio: T.priceHighRatio },
    rows
  };
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
    const earlier = series.slice(-(T.momentumMonths * 2), -T.momentumMonths);
    const recentAvg = recent.reduce((sum, entry) => sum + (entry.count || 0), 0) / recent.length;
    const earlierAvg = earlier.reduce((sum, entry) => sum + (entry.count || 0), 0) / earlier.length;
    const recentTotal = recentAvg * recent.length;
    const previousTotal = earlierAvg * earlier.length;
    const ratio = earlierAvg > 0 ? recentAvg / earlierAvg : recentAvg > 0 ? Infinity : 1;
    const changeRate = earlierAvg > 0 ? (recentAvg - earlierAvg) / earlierAvg : recentAvg > 0 ? 1 : 0;
    const row = {
      id: product.id,
      code: product.code,
      name: productLabel(product),
      ratio,
      changeRate,
      total,
      recentTotal,
      previousTotal,
      genre: primaryGenre(product)
    };
    if (recentTotal >= T.minMomentumWindowSales && ratio >= T.momentumRiseRatio) rising.push(row);
    else if (previousTotal >= T.minMomentumWindowSales && ratio <= T.momentumFallRatio) falling.push(row);
  });
  rising.sort((a, b) => (b.recentTotal - b.previousTotal) - (a.recentTotal - a.previousTotal));
  falling.sort((a, b) => (a.recentTotal - a.previousTotal) - (b.recentTotal - b.previousTotal));
  return { rising: rising.slice(0, T.maxListItems), falling: falling.slice(0, T.maxListItems) };
};

const percent = (value) => `${Math.round(value * 100)}%`;

// 分析結果 → 優先度付きの推奨アクション (根拠となる理論を明記する)
export const buildAdvisorActions = ({ abc, panelPerformance, balance, cannibalization, priceBands, momentum, profitability, textAnalysis, dataQuality }) => {
  const actions = [];

  if (abc.zeroSalesCount >= 3) {
    actions.push({
      priority: 'high',
      category: '品揃え',
      metric: `${abc.zeroSalesCount}商品`,
      title: `実績ゼロの商品を優先的に棚卸し`,
      detail: `${abc.zeroSales.slice(0, 3).map((p) => productLabel(p)).join('、')} などは誌面を使いながら販売数量がありません。差し替え・縮小 (1/16化)・カット候補として棚卸ししてください。`,
      theory: 'ABC分析: Cランク商品の整理は誌面の回転率 (スペース生産性) を直接引き上げる、小売の基本手法です。'
    });
  }

  if (panelPerformance.expansionCandidates.length > 0) {
    const names = panelPerformance.expansionCandidates.map((row) => row.name).join('、');
    actions.push({
      priority: 'mid',
      category: 'コマ配分',
      metric: `${panelPerformance.expansionCandidates.length}商品`,
      title: '総合実績の高い商品コマを広げる',
      detail: `${names} は、使用面積に対する数量・売上額・粗利額の総合実績が高い商品です。大コマ化や目立つ位置への移動を検討してください。`,
      theory: '総合スペース生産性: 利用できる実績指標を統合し、1/16コマ当たりの成果が高い商品へ誌面を配分します。'
    });
  }

  if (panelPerformance.reductionCandidates.length > 0) {
    const names = panelPerformance.reductionCandidates.map((row) => row.name).join('、');
    actions.push({
      priority: 'mid',
      category: 'コマ配分',
      metric: `${panelPerformance.reductionCandidates.length}商品`,
      title: '総合実績の低い大コマを見直す',
      detail: `${names} は、使用面積に対する総合実績が低い商品です。小コマ化し、実績の高い商品へ面積を振り替える候補です。`,
      theory: '総合スペース生産性: 数量だけでなく売上額と粗利額も加味し、低効率な大コマを見直します。'
    });
  }

  if (cannibalization.length > 0) {
    const pair = cannibalization[0];
    actions.push({
      priority: 'high',
      category: '重複',
      metric: `${cannibalization.length}組`,
      title: '役割が重なる商品を整理',
      detail: `最優先候補は「${pair.weak.name}」と「${pair.strong.name}」。同ジャンル「${pair.genre}」で類似度${percent(pair.similarity)}、${pair.basisLabel}は ${pair.weak.performanceValue.toLocaleString()} vs ${pair.strong.performanceValue.toLocaleString()} です。誌面テキストと価格差も考慮したうえで、弱い方の差し替えを検討してください。`,
      theory: 'カニバリゼーション回避と「選択のパラドックス」: 近似選択肢の並列は1商品あたりの購買率を下げます。差別化軸 (価格・サイズ・機能) を明確に分けるのが定石です。'
    });
  }

  balance.rows.filter((row) => row.status === 'under').slice(0, 2).forEach((row) => {
    actions.push({
      priority: 'mid',
      category: '誌面配分',
      metric: `効率 ${row.fairShare.toFixed(1)}×`,
      title: `「${row.genre}」は誌面が不足気味です`,
      detail: `総合実績シェア${percent(row.performanceShare)}に対し誌面面積は${percent(row.panelShare)}。実績に対して露出が少ないため、増コマ・大コマ化を検討できます。`,
      theory: 'フェアシェア理論: 露出シェアを数量・売上額・粗利額の総合実績シェアに近づけます。'
    });
  });
  balance.rows.filter((row) => row.status === 'over').slice(0, 2).forEach((row) => {
    actions.push({
      priority: 'mid',
      category: '誌面配分',
      metric: `効率 ${row.fairShare.toFixed(1)}×`,
      title: `「${row.genre}」は誌面過剰の可能性があります`,
      detail: `誌面面積${percent(row.panelShare)}に対し総合実績シェアは${percent(row.performanceShare)}。コマを絞り、実績の高いジャンルへ譲ることを検討してください。`,
      theory: '総合スペース生産性: 面積当たりの数量・売上額・粗利額が低い区画を見直します。'
    });
  });

  priceBands.rows.filter((row) => row.missing.length > 0).slice(0, 3).forEach((row) => {
    const bandLabel = { low: 'エントリー(低価格)', mid: 'ミドル', high: 'プレミアム(高価格)' };
    actions.push({
      priority: 'mid',
      category: '価格設計',
      metric: `${row.missing.length}帯不足`,
      title: `「${row.genre}」に${row.missing.map((band) => bandLabel[band]).join('・')}帯がありません`,
      detail: `ジャンル内の中央値 ${Math.round(row.median).toLocaleString()}円を基準に判定しました。価格の階段が欠けると予算の合わない顧客を逃すため、欠けた帯の商品追加を検討してください。`,
      theory: '価格ラダー / 松竹梅効果: 3価格帯を揃えると中位価格の選択率が上がることが知られています。'
    });
  });

  if (momentum.rising.length > 0) {
    actions.push({
      priority: 'info',
      category: '成長機会',
      metric: `${momentum.rising.length}商品`,
      title: `直近で伸びている商品が${momentum.rising.length}件あります`,
      detail: `${momentum.rising.slice(0, 3).map((row) => row.name).join('、')} などは直近3ヶ月がその前の3ヶ月を大きく上回っています。次号で大コマ化・前方配置・関連商品の併載を検討する価値があります。`,
      theory: 'モメンタム戦略: 伸びている商品への追加投資は、平均回帰前に需要を最大限取り込む定石です。'
    });
  }
  if (momentum.falling.length > 0) {
    actions.push({
      priority: 'info',
      category: '下降注意',
      metric: `${momentum.falling.length}商品`,
      title: `勢いが落ちている商品が${momentum.falling.length}件あります`,
      detail: `${momentum.falling.slice(0, 3).map((row) => row.name).join('、')} は直近3ヶ月が失速しています。季節要因か構造的な下降かを確認し、後者ならコマ縮小を。`,
      theory: 'プロダクトライフサイクル: 衰退期の商品は露出を段階的に縮小し、成長期の商品へ資源を移します。'
    });
  }

  if (abc.topShare >= 0.6) {
    actions.push({
      priority: 'info',
      category: '数量集中',
      metric: percent(abc.topShare),
      title: `販売数量の${percent(abc.topShare)}を上位10%の商品が占めています`,
      detail: '販売数量が上位商品へ集中しています。主力商品の欠品・廃盤に備えて代替候補を用意しつつ、数量上位商品は目立つ位置と十分なコマサイズを維持してください。',
      theory: 'パレートの法則 (80:20): 集中は効率的ですが、上位依存はリスク管理とセットで運用します。'
    });
  }
  if (profitability.lowMargin.length > 0) {
    actions.push({
      priority: 'high',
      category: '収益性',
      metric: `${profitability.lowMargin.length}商品`,
      title: '売上額は大きいが粗利率の低い商品を確認',
      detail: `${profitability.lowMargin.slice(0, 3).map((row) => `${row.name}（粗利率${percent(row.grossMargin)}）`).join('、')} は売上規模に対して粗利率が低めです。価格、仕入条件、掲載面積を併せて見直してください。`,
      theory: 'GMROIの考え方: 売上規模だけでなく粗利額と占有資源を同時に見ることで、利益に結びつく誌面配分を判断します。'
    });
  }
  if (textAnalysis.weakHighValue.length > 0) {
    actions.push({
      priority: 'mid',
      category: '誌面表現',
      metric: `${textAnalysis.weakHighValue.length}商品`,
      title: '実績上位商品の誌面テキストを補強',
      detail: `${textAnalysis.weakHighValue.slice(0, 3).map((row) => row.name).join('、')} は総合実績が高い一方、商品名・訴求・仕様などのテキスト情報が不足気味です。比較しやすい訴求へ整えてください。`,
      theory: '情報診断: 高価値商品のベネフィットと差別化軸を明示すると、誌面上での理解と比較を助けます。'
    });
  }
  if (dataQuality.quantityCoverage < 0.5) {
    actions.push({
      priority: 'info',
      category: 'データ品質',
      metric: percent(dataQuality.quantityCoverage),
      title: '販売数量データと照合できた商品が半数未満です',
      detail: '介援隊コードの記載漏れや販売数量CSVの期間ずれがあると、この分析の精度が下がります。コード整備と最新CSVの取り込みを先に行うと判断材料が揃います。',
      theory: 'データドリブン経営の前提は計測です。まず照合率を上げることが最も費用対効果の高い一手です。'
    });
  }

  const order = { high: 0, mid: 1, info: 2 };
  return actions.sort((a, b) => order[a.priority] - order[b.priority]).slice(0, T.maxActions);
};

export const buildAdvisorDataQuality = (products, vectorsById = null) => {
  const count = products.length;
  const ratio = (predicate) => count > 0 ? products.filter(predicate).length / count : 0;
  const quantityCoverage = ratio((product) => productHasMetric(product, PERFORMANCE_METRICS[0]));
  const salesAmountCoverage = ratio((product) => productHasMetric(product, PERFORMANCE_METRICS[1]));
  const grossProfitCoverage = ratio((product) => productHasMetric(product, PERFORMANCE_METRICS[2]));
  const priceCoverage = ratio((product) => toPrice(product.priceIncludingTax) !== null);
  const monthlyCoverage = ratio((product) => (product.monthlySales || []).length >= T.momentumMonths * 2);
  const textCoverage = ratio((product) => Number(product.catalogTextCompleteness) >= 0.5 || String(product.catalogText || product.sourceText || '').trim().length >= 40);
  const semanticCoverage = vectorsById
    ? ratio((product) => Array.isArray(vectorsById[product.id]) || ArrayBuffer.isView(vectorsById[product.id]))
    : 0;
  const score = Math.round((quantityCoverage * 0.25 + salesAmountCoverage * 0.2 + grossProfitCoverage * 0.25 + textCoverage * 0.15 + monthlyCoverage * 0.1 + priceCoverage * 0.05) * 100);
  return {
    score,
    level: score >= 80 ? 'high' : score >= 55 ? 'medium' : 'low',
    salesCoverage: quantityCoverage,
    quantityCoverage,
    salesAmountCoverage,
    grossProfitCoverage,
    textCoverage,
    priceCoverage,
    monthlyCoverage,
    semanticCoverage
  };
};

// エントリポイント: 掲載中の商品だけを対象に全セクションを計算する
export const buildCatalogAdvisorReport = ({ products = [], vectorsById = null } = {}) => {
  const placed = consolidateAdvisorProducts(products);
  const quantityProducts = placed.filter((product) => product.quantityMatched !== false && product.salesMatched !== false);

  const abc = buildAbcAnalysis(quantityProducts);
  const panelQuantity = buildPanelQuantityAnalysis(quantityProducts);
  const panelPerformance = buildPanelPerformanceAnalysis(placed);
  const balance = buildGenreBalance(placed);
  const cannibalization = buildCannibalizationPairs(placed, vectorsById);
  const priceBands = buildPriceBandCoverage(placed);
  const momentum = buildSalesMomentum(placed);
  const profitability = buildProfitabilityAnalysis(placed);
  const contributionRankings = buildContributionRankings(placed);
  const textAnalysis = buildCatalogTextAnalysis(placed, panelPerformance);
  const dataQuality = buildAdvisorDataQuality(placed, vectorsById);
  const actions = buildAdvisorActions({ abc, panelPerformance, balance, cannibalization, priceBands, momentum, profitability, textAnalysis, dataQuality });
  const priorityCounts = actions.reduce((counts, action) => ({
    ...counts,
    [action.priority]: counts[action.priority] + 1
  }), { high: 0, mid: 0, info: 0 });

  return {
    generatedAt: new Date().toISOString(),
    summary: {
      placedCount: placed.length,
      totalSales: abc.totalSales,
      totalQuantity: abc.totalSales,
      totalSalesAmount: balance.totalSalesAmount,
      totalGrossProfitAmount: balance.totalGrossProfitAmount,
      grossMargin: profitability.grossMargin,
      salesCoverage: dataQuality.quantityCoverage,
      quantityCoverage: dataQuality.quantityCoverage,
      salesAmountCoverage: dataQuality.salesAmountCoverage,
      grossProfitCoverage: dataQuality.grossProfitCoverage,
      textCoverage: dataQuality.textCoverage,
      availableMetrics: panelPerformance.availableMetrics,
      topShare: abc.topShare,
      genreCount: balance.rows.length,
      placementCount: placed.reduce((sum, product) => sum + product.assignments.length, 0),
      spaceUnits: balance.totalSpaceUnits,
      usedEmbeddings: !!vectorsById,
      dataQualityScore: dataQuality.score,
      priorityCounts
    },
    abc,
    panelQuantity,
    panelPerformance,
    balance,
    cannibalization,
    priceBands,
    momentum,
    profitability,
    contributionRankings,
    textAnalysis,
    dataQuality,
    actions
  };
};
