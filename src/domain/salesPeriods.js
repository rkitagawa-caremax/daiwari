export const SALES_PERIOD_CURRENT = 'current';
export const SALES_PERIOD_PREVIOUS = 'previous';
export const SALES_PERIOD_TWO_PREVIOUS = 'twoPrevious';

export const SALES_PERIOD_OPTIONS = Object.freeze([
  Object.freeze({
    id: SALES_PERIOD_CURRENT,
    label: '今期',
    description: '現在進行中の年度',
    chunkCollectionId: 'salesDataChunks',
    metaDocumentId: 'salesDataMeta',
    localDataKey: 'salesData',
    localMetaKey: 'salesDataMeta'
  }),
  Object.freeze({
    id: SALES_PERIOD_PREVIOUS,
    label: '前期',
    description: '1年前の年度',
    chunkCollectionId: 'salesDataChunksPrevious',
    metaDocumentId: 'salesDataMetaPrevious',
    localDataKey: 'salesData:previous',
    localMetaKey: 'salesDataMeta:previous'
  }),
  Object.freeze({
    id: SALES_PERIOD_TWO_PREVIOUS,
    label: '前々期',
    description: '2年前の年度',
    chunkCollectionId: 'salesDataChunksTwoPrevious',
    metaDocumentId: 'salesDataMetaTwoPrevious',
    localDataKey: 'salesData:twoPrevious',
    localMetaKey: 'salesDataMeta:twoPrevious'
  })
]);

const salesPeriodById = new Map(SALES_PERIOD_OPTIONS.map((period) => [period.id, period]));

export const getSalesPeriodDefinition = (periodId) => (
  salesPeriodById.get(periodId) || salesPeriodById.get(SALES_PERIOD_CURRENT)
);

export const getSalesPeriodCacheKey = (baseKey, periodId) => (
  periodId === SALES_PERIOD_CURRENT ? baseKey : `${baseKey}:${getSalesPeriodDefinition(periodId).id}`
);

export const createEmptySalesPeriodMeta = () => Object.fromEntries(
  SALES_PERIOD_OPTIONS.map((period) => [period.id, null])
);

export const getSalesDataMonthLabels = (salesData) => {
  for (const items of Object.values(salesData || {})) {
    if (!Array.isArray(items)) continue;
    const labels = items.find((item) => Array.isArray(item?.monthlyLabels))?.monthlyLabels;
    if (labels?.length) return labels.filter(Boolean);
  }
  return [];
};

export const buildSalesPeriodMeta = ({ salesData, fileName = '', updatedAt = new Date() } = {}) => ({
  fileName,
  updatedAt,
  totalItems: Object.keys(salesData || {}).length,
  totalRows: Object.values(salesData || {}).reduce(
    (total, items) => total + (Array.isArray(items) ? items.length : 0),
    0
  ),
  monthLabels: getSalesDataMonthLabels(salesData)
});

