import React, { useId, useMemo } from 'react';

const CHART_LEFT = 10;
const CHART_RIGHT = 98;
const CHART_TOP = 3;
const CHART_BOTTOM = 45;

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
    const linePoints = points.map((point) => `${point.x},${point.y}`).join(' ');
    const areaPoints = points.length > 0
      ? `${points[0].x},${CHART_BOTTOM} ${linePoints} ${points.at(-1).x},${CHART_BOTTOM}`
      : '';
    return { maximum, points, linePoints, areaPoints };
  }, [series]);

  if (series.length === 0) return null;
  const middleIndex = Math.floor((series.length - 1) / 2);
  const labelIndexes = [...new Set([0, middleIndex, series.length - 1])];

  return (
    <div className="flex min-h-0 flex-1 flex-col" aria-label="月別売上折れ線グラフ">
      <div className="mb-0.5 flex items-center justify-between text-[8px] font-bold leading-none text-cyan-50">
        <span>月別推移</span>
        <span className="font-mono">最大 {chart.maximum.toLocaleString()}</span>
      </div>
      <svg
        viewBox="0 0 100 52"
        preserveAspectRatio="none"
        className="min-h-0 w-full flex-1 overflow-visible"
        role="img"
        aria-label={series.map((entry) => `${entry.label} ${entry.count}`).join('、')}
      >
        <defs>
          <linearGradient id={gradientId} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#22d3ee" stopOpacity="0.94" />
            <stop offset="58%" stopColor="#38bdf8" stopOpacity="0.78" />
            <stop offset="100%" stopColor="#bae6fd" stopOpacity="0.34" />
          </linearGradient>
        </defs>
        {[CHART_TOP, CHART_TOP + ((CHART_BOTTOM - CHART_TOP) / 3), CHART_TOP + (((CHART_BOTTOM - CHART_TOP) * 2) / 3), CHART_BOTTOM].map((y) => (
          <line key={y} x1={CHART_LEFT} y1={y} x2={CHART_RIGHT} y2={y} stroke="rgba(255,255,255,0.24)" strokeWidth="0.65" vectorEffect="non-scaling-stroke" />
        ))}
        <text x="0" y={CHART_TOP + 2} fill="rgba(255,255,255,0.62)" fontSize="4.5" fontWeight="700">{chart.maximum}</text>
        <text x="0" y={((CHART_TOP + CHART_BOTTOM) / 2) + 2} fill="rgba(255,255,255,0.5)" fontSize="4.5" fontWeight="700">{Math.round(chart.maximum / 2)}</text>
        <text x="2" y={CHART_BOTTOM + 1} fill="rgba(255,255,255,0.62)" fontSize="4.5" fontWeight="700">0</text>
        <polygon points={chart.areaPoints} fill={`url(#${gradientId})`} />
        <polyline points={chart.linePoints} fill="none" stroke="#22d3ee" strokeWidth="2" strokeLinejoin="miter" strokeLinecap="square" vectorEffect="non-scaling-stroke" />
        {labelIndexes.map((index) => {
          const point = chart.points[index];
          if (!point) return null;
          return (
            <text key={`label-${index}`} x={point.x} y="50" textAnchor={index === 0 ? 'start' : index === series.length - 1 ? 'end' : 'middle'} fill="rgba(255,255,255,0.78)" fontSize="4.5" fontWeight="700">
              {point.label}
            </text>
          );
        })}
      </svg>
    </div>
  );
});

MonthlySalesChart.displayName = 'MonthlySalesChart';

export default MonthlySalesChart;
