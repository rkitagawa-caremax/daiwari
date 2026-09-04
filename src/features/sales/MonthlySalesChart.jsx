import React, { useId, useMemo } from 'react';

const CHART_LEFT = 4;
const CHART_RIGHT = 96;
const CHART_TOP = 4;
const CHART_BOTTOM = 44;

const round2 = (value) => Math.round(value * 100) / 100;
const clampY = (value) => Math.min(CHART_BOTTOM, Math.max(CHART_TOP - 1.5, value));

// Catmull-Rom をベジェに変換したなめらかな折れ線。
// 制御点の y は描画域に収め、スパイクでベースラインを突き抜けないようにする。
const buildSmoothPath = (points) => {
  if (points.length === 0) return '';
  let path = `M ${round2(points[0].x)},${round2(points[0].y)}`;
  for (let index = 0; index < points.length - 1; index++) {
    const previous = points[index - 1] || points[index];
    const current = points[index];
    const next = points[index + 1];
    const afterNext = points[index + 2] || next;
    const c1x = current.x + (next.x - previous.x) / 6;
    const c1y = clampY(current.y + (next.y - previous.y) / 6);
    const c2x = next.x - (afterNext.x - current.x) / 6;
    const c2y = clampY(next.y - (afterNext.y - current.y) / 6);
    path += ` C ${round2(c1x)},${round2(c1y)} ${round2(c2x)},${round2(c2y)} ${round2(next.x)},${round2(next.y)}`;
  }
  return path;
};

const MonthlySalesChart = React.memo(({ series = [] }) => {
  const gradientId = `monthly-sales-area-${useId().replace(/:/g, '')}`;
  const chart = useMemo(() => {
    const values = series.map((entry) => Math.max(0, Number(entry?.count) || 0));
    const maximum = Math.max(1, ...values);
    const step = values.length > 1 ? (CHART_RIGHT - CHART_LEFT) / (values.length - 1) : 0;
    const points = values.map((value, index) => ({
      x: values.length === 1 ? 50 : CHART_LEFT + step * index,
      y: CHART_BOTTOM - ((value / maximum) * (CHART_BOTTOM - CHART_TOP)),
      value,
      label: series[index]?.label || ''
    }));
    const linePath = buildSmoothPath(points);
    const areaPath = points.length > 0
      ? `${linePath} L ${round2(points.at(-1).x)},${CHART_BOTTOM} L ${round2(points[0].x)},${CHART_BOTTOM} Z`
      : '';
    // 最大値の月 (同数の場合は最初の月)。全月 0 のときは無し
    const maxValue = Math.max(0, ...values);
    const maxIndex = maxValue > 0 ? values.indexOf(maxValue) : -1;
    return {
      maximum,
      maximumLabel: maxIndex >= 0 ? (series[maxIndex]?.label || '') : '',
      maxPoint: maxIndex >= 0 ? points[maxIndex] : null,
      points,
      linePath,
      areaPath
    };
  }, [series]);

  if (series.length === 0) return null;
  const middleIndex = Math.floor((series.length - 1) / 2);
  const labelIndexes = [...new Set([0, middleIndex, series.length - 1])];

  return (
    <div className="flex min-h-0 flex-1 flex-col" aria-label="月別売上折れ線グラフ">
      <div className="mb-1 flex items-baseline justify-between gap-1 leading-none text-cyan-50">
        <span className="text-[11px] font-bold">月別推移</span>
        <span className="flex items-baseline gap-1 whitespace-nowrap">
          <span className="text-[11px] font-bold">最大</span>
          <span className="font-mono text-2xl font-black leading-none tracking-tight text-cyan-200">{chart.maximum.toLocaleString()}</span>
          {chart.maximumLabel && <span className="text-[11px] font-bold">（{chart.maximumLabel}）</span>}
        </span>
      </div>
      <svg
        viewBox="0 0 100 46"
        preserveAspectRatio="none"
        className="min-h-0 w-full flex-1 overflow-visible"
        role="img"
        aria-label={series.map((entry) => `${entry.label} ${entry.count}`).join('、')}
      >
        <defs>
          <linearGradient id={gradientId} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#22d3ee" stopOpacity="0.5" />
            <stop offset="100%" stopColor="#22d3ee" stopOpacity="0" />
          </linearGradient>
        </defs>

        {/* 基準線は最大値 (破線) とベースラインの 2 本だけに絞る */}
        <line x1={CHART_LEFT} y1={CHART_TOP} x2={CHART_RIGHT} y2={CHART_TOP} stroke="rgba(255,255,255,0.18)" strokeWidth="0.8" strokeDasharray="2 3" vectorEffect="non-scaling-stroke" />
        <line x1={CHART_LEFT} y1={CHART_BOTTOM} x2={CHART_RIGHT} y2={CHART_BOTTOM} stroke="rgba(255,255,255,0.3)" strokeWidth="0.8" vectorEffect="non-scaling-stroke" />

        <path d={chart.areaPath} fill={`url(#${gradientId})`} />
        <path d={chart.linePath} fill="none" stroke="#67e8f9" strokeWidth="1.75" strokeLinejoin="round" strokeLinecap="round" vectorEffect="non-scaling-stroke" />

        {/* 最大値の月に打点する。長さ 0 のパス + 丸キャップなので縦横比が歪んでも真円のまま */}
        {chart.maxPoint && (
          <>
            <path d={`M ${round2(chart.maxPoint.x)},${round2(chart.maxPoint.y)} l 0.001,0`} stroke="rgba(34,211,238,0.35)" strokeWidth="8" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
            <path d={`M ${round2(chart.maxPoint.x)},${round2(chart.maxPoint.y)} l 0.001,0`} stroke="#ffffff" strokeWidth="3.5" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
          </>
        )}

      </svg>
      {/* 月ラベルは SVG の外に出す。SVG 内だと縦横比の引き伸ばしで文字が歪んで読みにくい */}
      <div className="mt-1 flex items-center justify-between text-[11px] font-bold leading-none text-cyan-300">
        {labelIndexes.map((index) => (
          <span key={`label-${index}`}>{series[index]?.label || ''}</span>
        ))}
      </div>
    </div>
  );
});

MonthlySalesChart.displayName = 'MonthlySalesChart';

export default MonthlySalesChart;
