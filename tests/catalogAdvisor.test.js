import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildAbcAnalysis,
  buildCannibalizationPairs,
  buildCatalogAdvisorReport,
  buildGenreBalance,
  buildPanelPerformanceAnalysis,
  buildProfitabilityAnalysis,
  buildPanelQuantityAnalysis,
  buildPriceBandCoverage,
  buildSalesMomentum,
  consolidateAdvisorProducts
} from '../src/domain/catalogAdvisor.js';

const months = (counts) => counts.map((count, index) => ({ label: `${index + 1}月`, count }));

const makeProduct = (overrides = {}) => ({
  id: overrides.id || overrides.code || 'p',
  code: overrides.code || 'E0001',
  name: overrides.name || `商品${overrides.code || ''}`,
  itemNumber: '',
  priceIncludingTax: overrides.priceIncludingTax ?? '',
  lifecycleStatus: overrides.lifecycleStatus || '',
  searchText: overrides.searchText || `${overrides.name || ''} ${overrides.code || ''}`,
  assignments: overrides.assignments ?? [{ sheetId: 's1', pageNumber: 1, panelIndex: 0, genre: overrides.genre || '食事関連' }],
  salesCount: overrides.salesCount ?? 0,
  monthlySales: overrides.monthlySales ?? [],
  ...overrides
});

test('buildAbcAnalysis ranks by cumulative share and lists zero-sales products', () => {
  const products = [
    makeProduct({ id: 'a', code: 'E0001', salesCount: 800 }),
    makeProduct({ id: 'b', code: 'E0002', salesCount: 150 }),
    makeProduct({ id: 'c', code: 'E0003', salesCount: 50 }),
    makeProduct({ id: 'd', code: 'E0004', salesCount: 0 }),
    makeProduct({ id: 'e', code: 'E0005', salesCount: 0, lifecycleStatus: '廃盤' })
  ];
  const abc = buildAbcAnalysis(products);
  assert.equal(abc.totalSales, 1000);
  assert.equal(abc.rankByProductId.a, 'A');
  assert.equal(abc.rankByProductId.b, 'B');
  assert.equal(abc.rankByProductId.c, 'C');
  assert.equal(abc.rankByProductId.d, 'C');
  // 廃盤は死に筋リストから除外する
  assert.deepEqual(abc.zeroSales.map((p) => p.id), ['d']);
  assert.ok(abc.topShare > 0.7, '上位10% (=1商品) が 800/1000 を占める');
});

test('buildAbcAnalysis keeps the product crossing the A threshold in rank A', () => {
  const products = [
    makeProduct({ id: 'dominant', salesCount: 900 }),
    makeProduct({ id: 'second', salesCount: 60 }),
    makeProduct({ id: 'third', salesCount: 40 })
  ];
  const abc = buildAbcAnalysis(products);
  assert.equal(abc.rankByProductId.dominant, 'A');
  assert.equal(abc.rankByProductId.second, 'B');
  assert.equal(abc.rankByProductId.third, 'C');
});

test('buildGenreBalance flags over- and under-allocated genres by fair share', () => {
  const products = [
    // 入浴関連: コマ8つで売上わずか → over
    ...Array.from({ length: 8 }, (_, i) => makeProduct({ id: `bath-${i}`, code: `S00${i}`, genre: '入浴関連', salesCount: 1 })),
    // 食事関連: コマ6つで売上の大半 → under
    ...Array.from({ length: 6 }, (_, i) => makeProduct({ id: `food-${i}`, code: `E00${i}`, genre: '食事関連', salesCount: 200 }))
  ];
  const balance = buildGenreBalance(products);
  const bath = balance.rows.find((row) => row.genre === '入浴関連');
  const food = balance.rows.find((row) => row.genre === '食事関連');
  assert.equal(bath.status, 'over');
  assert.equal(food.status, 'under');
  assert.ok(Math.abs(balance.rows.reduce((sum, row) => sum + row.panelShare, 0) - 1) < 1e-9);
});

test('buildPanelQuantityAnalysis compares total quantity and quantity per space unit', () => {
  const compactSeller = makeProduct({
    id: 'compact',
    salesCount: 120,
    salesMatched: true,
    assignments: [{ sheetId: 's1', panelIndex: 0, genre: '食事関連', rowSpan: 1, colSpan: 1 }]
  });
  const largeSlowSeller = makeProduct({
    id: 'large',
    salesCount: 8,
    salesMatched: true,
    assignments: [{ sheetId: 's1', panelIndex: 1, genre: '食事関連', rowSpan: 2, colSpan: 2 }]
  });
  const analysis = buildPanelQuantityAnalysis([compactSeller, largeSlowSeller]);
  assert.equal(analysis.highest[0].id, 'compact');
  assert.equal(analysis.highest[0].quantityPerSpace, 120);
  assert.equal(analysis.lowest[0].id, 'large');
  assert.equal(analysis.lowest[0].quantityPerSpace, 2);
  assert.deepEqual(analysis.expansionCandidates.map((row) => row.id), ['compact']);
  assert.deepEqual(analysis.reductionCandidates.map((row) => row.id), ['large']);
});

test('panel performance combines quantity, sales amount and gross profit without penalizing unavailable metrics', () => {
  const profitLeader = makeProduct({
    id: 'profit', salesCount: 20, quantityMatched: true,
    salesAmount: 100000, salesAmountMatched: true,
    grossProfitAmount: 50000, grossProfitMatched: true
  });
  const quantityLeader = makeProduct({
    id: 'quantity', salesCount: 100, quantityMatched: true,
    salesAmount: 100000, salesAmountMatched: true,
    grossProfitAmount: 10000, grossProfitMatched: true
  });
  const analysis = buildPanelPerformanceAnalysis([profitLeader, quantityLeader]);
  assert.deepEqual(analysis.availableMetrics, ['quantity', 'salesAmount', 'grossProfitAmount']);
  assert.equal(analysis.highest[0].id, 'profit', '粗利額を最重視した総合評価になる');

  const quantityOnly = buildPanelPerformanceAnalysis([
    makeProduct({ id: 'q1', salesCount: 10 }),
    makeProduct({ id: 'q2', salesCount: 20 })
  ]);
  assert.deepEqual(quantityOnly.availableMetrics, ['quantity']);
  assert.equal(quantityOnly.highest[0].id, 'q2');
});

test('profitability analysis detects high-sales low-margin products', () => {
  const result = buildProfitabilityAnalysis([
    makeProduct({ id: 'low', salesAmount: 200000, salesAmountMatched: true, grossProfitAmount: 10000, grossProfitMatched: true }),
    makeProduct({ id: 'healthy', salesAmount: 100000, salesAmountMatched: true, grossProfitAmount: 50000, grossProfitMatched: true })
  ]);
  assert.equal(result.totalSalesAmount, 300000);
  assert.deepEqual(result.lowMargin.map((row) => row.id), ['low']);
});

test('buildGenreBalance weights large panels and splits sales across assigned genres', () => {
  const product = makeProduct({
    id: 'multi',
    salesCount: 100,
    assignments: [
      { sheetId: 's1', panelIndex: 0, genre: '食事関連', rowSpan: 2, colSpan: 2 },
      { sheetId: 's2', panelIndex: 0, genre: '入浴関連', rowSpan: 1, colSpan: 1 }
    ]
  });
  const balance = buildGenreBalance([product]);
  const food = balance.rows.find((row) => row.genre === '食事関連');
  const bath = balance.rows.find((row) => row.genre === '入浴関連');
  assert.equal(balance.totalSpaceUnits, 5);
  assert.equal(food.spaceUnits, 4);
  assert.equal(food.sales, 80);
  assert.equal(bath.sales, 20);
});

test('consolidateAdvisorProducts prevents duplicate SKU sales and placements', () => {
  const products = [
    makeProduct({ id: 'first', code: 'E1000', salesCount: 25 }),
    makeProduct({
      id: 'duplicate',
      code: 'E1000',
      salesCount: 25,
      assignments: [{ sheetId: 's2', pageNumber: 2, panelIndex: 3, genre: '食事関連' }]
    })
  ];
  const consolidated = consolidateAdvisorProducts(products);
  assert.equal(consolidated.length, 1);
  assert.equal(consolidated[0].salesCount, 25);
  assert.equal(consolidated[0].assignments.length, 2);
});

test('buildCannibalizationPairs uses embeddings when available and flags weak twins', () => {
  const strong = makeProduct({ id: 'strong', code: 'E0100', salesCount: 100, genre: '食事関連' });
  const weakTwin = makeProduct({ id: 'weak', code: 'E0101', salesCount: 5, genre: '食事関連' });
  const unrelated = makeProduct({ id: 'other', code: 'E0200', salesCount: 90, genre: '食事関連' });
  const vectors = {
    strong: [1, 0, 0],
    weak: [0.999, 0.04, 0],
    other: [0, 1, 0]
  };
  const pairs = buildCannibalizationPairs([strong, weakTwin, unrelated], vectors);
  assert.equal(pairs.length, 1);
  assert.equal(pairs[0].method, 'embedding');
  assert.equal(pairs[0].strong.id, 'strong');
  assert.equal(pairs[0].weak.id, 'weak');

  // 実績が拮抗しているペアはカニバリ扱いしない
  const evenTwin = { ...weakTwin, salesCount: 80 };
  assert.equal(buildCannibalizationPairs([strong, evenTwin], vectors).length, 0);
});

test('buildCannibalizationPairs falls back to lexical similarity without vectors', () => {
  const strong = makeProduct({
    id: 'ls', code: 'E0300', salesCount: 60, genre: '食事関連',
    searchText: 'やさしくラクケア まるで果物のようなゼリー りんご味 低カロリー 80g'
  });
  const weak = makeProduct({
    id: 'lw', code: 'E0301', salesCount: 2, genre: '食事関連',
    searchText: 'やさしくラクケア まるで果物のようなゼリー もも味 低カロリー 80g'
  });
  const different = makeProduct({
    id: 'ld', code: 'E0302', salesCount: 3, genre: '食事関連',
    searchText: '車いす用クッション 体圧分散 撥水カバー'
  });
  const pairs = buildCannibalizationPairs([strong, weak, different], null);
  assert.equal(pairs.length, 1);
  assert.equal(pairs[0].method, 'lexical');
  assert.equal(pairs[0].weak.id, 'lw');
});

test('buildPriceBandCoverage reports genres missing a price band', () => {
  const products = [
    // 別ジャンルの価格水準は食事関連の判定に影響しない
    ...[500, 800, 1000, 1500, 2000, 3000, 5000, 8000, 12000].map((price, i) => (
      makeProduct({ id: `bg-${i}`, code: `W0${i}`, genre: '歩行関連', priceIncludingTax: price })
    )),
    // 食事関連は中央値付近だけで、エントリーとプレミアムがない
    ...[300, 320, 340, 360, 380].map((price, i) => (
      makeProduct({ id: `fd-${i}`, code: `E05${i}`, genre: '食事関連', priceIncludingTax: price })
    ))
  ];
  const coverage = buildPriceBandCoverage(products);
  assert.ok(coverage.bands);
  const food = coverage.rows.find((row) => row.genre === '食事関連');
  assert.ok(food.missing.includes('low'));
  assert.ok(food.missing.includes('high'));
  assert.equal(food.mid, 5);
  assert.equal(food.median, 340);
});

test('buildSalesMomentum separates rising and falling products', () => {
  const rising = makeProduct({
    id: 'up', code: 'E0400', salesCount: 60,
    monthlySales: months([2, 2, 2, 2, 2, 2, 2, 2, 2, 14, 14, 14])
  });
  const falling = makeProduct({
    id: 'down', code: 'E0401', salesCount: 60,
    monthlySales: months([10, 10, 10, 10, 10, 10, 10, 10, 10, 1, 1, 1])
  });
  const flat = makeProduct({
    id: 'flat', code: 'E0402', salesCount: 24,
    monthlySales: months([2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])
  });
  const momentum = buildSalesMomentum([rising, falling, flat]);
  assert.deepEqual(momentum.rising.map((row) => row.id), ['up']);
  assert.deepEqual(momentum.falling.map((row) => row.id), ['down']);
});

test('buildCatalogAdvisorReport produces prioritized actions with theory notes', () => {
  const products = [
    makeProduct({ id: 'a', code: 'E0001', salesCount: 900, genre: '食事関連' }),
    makeProduct({ id: 'b', code: 'E0002', salesCount: 60, genre: '食事関連' }),
    makeProduct({ id: 'z1', code: 'E0003', salesCount: 0, genre: '入浴関連' }),
    makeProduct({ id: 'z2', code: 'E0004', salesCount: 0, genre: '入浴関連' }),
    makeProduct({ id: 'z3', code: 'E0005', salesCount: 0, genre: '入浴関連' }),
    // 未配置の商品は分析対象に含めない
    makeProduct({ id: 'lib', code: 'E0006', salesCount: 500, assignments: [] })
  ];
  const report = buildCatalogAdvisorReport({ products });
  assert.equal(report.summary.placedCount, 5);
  assert.equal(report.summary.usedEmbeddings, false);
  assert.equal(report.abc.totalSales, 960, '未配置の500は合算されない');
  assert.ok(report.actions.length > 0);
  assert.ok(report.actions.every((action) => action.title && action.detail && action.theory));
  const priorities = report.actions.map((action) => action.priority);
  const firstMid = priorities.indexOf('mid');
  const lastHigh = priorities.lastIndexOf('high');
  if (firstMid !== -1 && lastHigh !== -1) assert.ok(lastHigh < firstMid, 'high が mid より先に並ぶ');
  // 死に筋3件 → high アクションが立つ
  assert.ok(report.actions.some((action) => action.priority === 'high' && action.title.includes('実績ゼロ')));
});
