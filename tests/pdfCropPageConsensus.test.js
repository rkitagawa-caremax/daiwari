import test from 'node:test';
import assert from 'node:assert/strict';

import {
  selectPdfCropConsensusEdge,
  stabilizePdfCropPageRects
} from '../src/domain/pdfCropPageConsensus.js';

const GRID = { left: 0.1, top: 0.06, cellWidth: 0.2, cellHeight: 0.22 };
const row = (id, xPos, yPos) => ({ id, xPos, yPos, colSpan: 1, rowSpan: 1 });

test('selectPdfCropConsensusEdge uses the densest cluster instead of an isolated false edge', () => {
  const selected = selectPdfCropConsensusEdge([0.298, 0.301, 0.302, 0.325], 0.3, 0.2);
  assert.equal(selected.count, 3);
  assert.ok(Math.abs(selected.value - 0.301) < 0.002);
});

test('stabilizePdfCropPageRects shares a reliable column edge with rows that missed it', () => {
  const entries = [
    {
      row: row('a', 1, 1),
      rect: { x: 0.102, y: 0.061, width: 0.196, height: 0.218 },
      snappedEdges: { left: true, right: true, top: true, bottom: true },
      snappedCount: 4
    },
    {
      row: row('b', 1, 2),
      rect: { x: 0.101, y: 0.282, width: 0.199, height: 0.216 },
      snappedEdges: { left: true, right: false, top: true, bottom: true },
      snappedCount: 3
    },
    {
      row: row('c', 1, 3),
      rect: { x: 0.103, y: 0.501, width: 0.196, height: 0.218 },
      snappedEdges: { left: true, right: true, top: true, bottom: true },
      snappedCount: 4
    },
    {
      row: row('d', 1, 4),
      rect: { x: 0.102, y: 0.721, width: 0.226, height: 0.218 },
      snappedEdges: { left: true, right: true, top: true, bottom: true },
      snappedCount: 4
    }
  ];
  const stabilized = stabilizePdfCropPageRects({ entries, grid: GRID });
  const sharedRight = stabilized[0].rect.x + stabilized[0].rect.width;
  assert.ok(Math.abs(sharedRight - 0.2985) < 1e-6);
  assert.ok(Math.abs(stabilized[1].rect.x + stabilized[1].rect.width - sharedRight) < 1e-6);
  assert.ok(Math.abs(stabilized[2].rect.x + stabilized[2].rect.width - sharedRight) < 1e-6);
  assert.equal(stabilized[1].consensusEdges.right, true);
});

test('stabilizePdfCropPageRects preserves a manual frame', () => {
  const manual = {
    row: row('manual', 1, 1),
    rect: { x: 0.12, y: 0.08, width: 0.17, height: 0.19 },
    isManual: true
  };
  assert.equal(stabilizePdfCropPageRects({ entries: [manual], grid: GRID })[0], manual);
});
