// 同じページのコマ枠を、4x4 グリッド上の共通境界ごとに安定化する。
// 各コマを単独で画像解析すると、商品写真内の白帯や細線を境界と誤認することがある。
// 同じ論理線を使う複数コマの候補を集め、最も密集した候補群の中央値を採用することで、
// 外れ値を捨てつつ、境界を検出できなかったコマにも同じ行・列の結果を反映できる。

export const DEFAULT_PDF_CROP_CONSENSUS_OPTIONS = Object.freeze({
  searchToleranceRatio: 0.16,
  clusterToleranceRatio: 0.025,
  sizeToleranceRatio: 0.32
});

const median = (values) => {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};

export const selectPdfCropConsensusEdge = (
  values = [],
  expected,
  cellSize,
  options = DEFAULT_PDF_CROP_CONSENSUS_OPTIONS
) => {
  if (!Number.isFinite(expected) || !Number.isFinite(cellSize) || cellSize <= 0) return null;
  const searchTolerance = cellSize * options.searchToleranceRatio;
  const clusterTolerance = cellSize * options.clusterToleranceRatio;
  const candidates = values.filter((value) => (
    Number.isFinite(value) && Math.abs(value - expected) <= searchTolerance
  ));
  if (candidates.length === 0) return null;

  const clusters = candidates.map((center) => {
    const members = candidates.filter((value) => Math.abs(value - center) <= clusterTolerance);
    const value = median(members);
    const spread = Math.max(...members.map((member) => Math.abs(member - value)), 0);
    return { value, count: members.length, spread, distance: Math.abs(value - expected) };
  });
  clusters.sort((left, right) => (
    right.count - left.count
    || left.spread - right.spread
    || left.distance - right.distance
  ));
  return clusters[0];
};

const addCandidate = (groups, key, value) => {
  if (!Number.isFinite(value)) return;
  if (!groups.has(key)) groups.set(key, []);
  groups.get(key).push(value);
};

const getLineKey = (axis, side, index) => `${axis}:${side}:${index}`;

export const stabilizePdfCropPageRects = ({ entries = [], grid, options = {} } = {}) => {
  if (!grid || !Array.isArray(entries) || entries.length === 0) return entries;
  const resolved = { ...DEFAULT_PDF_CROP_CONSENSUS_OPTIONS, ...options };
  const groups = new Map();

  entries.forEach((entry) => {
    const { row, rect, snappedEdges = {} } = entry || {};
    if (!row || !rect || entry.isManual) return;
    const leftLine = row.xPos - 1;
    const rightLine = leftLine + row.colSpan;
    const topLine = row.yPos - 1;
    const bottomLine = topLine + row.rowSpan;
    if (snappedEdges.left) addCandidate(groups, getLineKey('x', 'start', leftLine), rect.x);
    if (snappedEdges.right) addCandidate(groups, getLineKey('x', 'end', rightLine), rect.x + rect.width);
    if (snappedEdges.top) addCandidate(groups, getLineKey('y', 'start', topLine), rect.y);
    if (snappedEdges.bottom) addCandidate(groups, getLineKey('y', 'end', bottomLine), rect.y + rect.height);
  });

  const consensus = new Map();
  groups.forEach((values, key) => {
    const [axis, , rawIndex] = key.split(':');
    const index = Number.parseInt(rawIndex, 10);
    const cellSize = axis === 'x' ? grid.cellWidth : grid.cellHeight;
    const origin = axis === 'x' ? grid.left : grid.top;
    const selected = selectPdfCropConsensusEdge(values, origin + index * cellSize, cellSize, resolved);
    if (selected) consensus.set(key, selected);
  });

  return entries.map((entry) => {
    const { row, rect, snappedEdges = {} } = entry || {};
    if (!row || !rect || entry.isManual) return entry;
    const leftLine = row.xPos - 1;
    const rightLine = leftLine + row.colSpan;
    const topLine = row.yPos - 1;
    const bottomLine = topLine + row.rowSpan;
    // 1件だけの候補は、そのコマ自身が検出した辺にだけ使う。
    // 未検出のコマへ共有するのは、2コマ以上が同じ境界を支持した場合に限る。
    const usable = (candidate, ownEdge) => candidate && (ownEdge || candidate.count >= 2) ? candidate : null;
    const leftConsensus = usable(consensus.get(getLineKey('x', 'start', leftLine)), snappedEdges.left);
    const rightConsensus = usable(consensus.get(getLineKey('x', 'end', rightLine)), snappedEdges.right);
    const topConsensus = usable(consensus.get(getLineKey('y', 'start', topLine)), snappedEdges.top);
    const bottomConsensus = usable(consensus.get(getLineKey('y', 'end', bottomLine)), snappedEdges.bottom);
    const left = leftConsensus?.value ?? rect.x;
    const right = rightConsensus?.value ?? (rect.x + rect.width);
    const top = topConsensus?.value ?? rect.y;
    const bottom = bottomConsensus?.value ?? (rect.y + rect.height);
    const expectedWidth = row.colSpan * grid.cellWidth;
    const expectedHeight = row.rowSpan * grid.cellHeight;
    const width = right - left;
    const height = bottom - top;
    const plausible = width > 0
      && height > 0
      && Math.abs(width - expectedWidth) <= expectedWidth * resolved.sizeToleranceRatio
      && Math.abs(height - expectedHeight) <= expectedHeight * resolved.sizeToleranceRatio;
    if (!plausible) return entry;

    const consensusEdges = {
      left: !!leftConsensus,
      right: !!rightConsensus,
      top: !!topConsensus,
      bottom: !!bottomConsensus
    };
    return {
      ...entry,
      rect: { x: left, y: top, width, height },
      consensusEdges,
      consensusCount: Object.values(consensusEdges).filter(Boolean).length,
      snappedCount: Math.max(entry.snappedCount || 0, Object.values(consensusEdges).filter(Boolean).length)
    };
  });
};
