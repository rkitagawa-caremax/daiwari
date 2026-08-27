import test from 'node:test';
import assert from 'node:assert/strict';

import {
  addWorkActionDelta,
  aggregateWorkLogRecords,
  applyWorkLogDeltaToRecord,
  createEmptyWorkLogDelta,
  formatWorkDuration,
  getJstDateKey,
  resolveWorkActionId
} from '../src/domain/workActivity.js';

test('work action classifier only returns fixed non-sensitive categories', () => {
  assert.equal(resolveWorkActionId({ explicitAction: 'pdf_export' }), 'pdf_export');
  assert.equal(resolveWorkActionId({ explicitAction: 'unknown', text: 'PDF出力' }), 'pdf_export');
  assert.equal(resolveWorkActionId({ text: '商品E1931を編集', isPanel: true }), 'panel_edit');
  assert.equal(resolveWorkActionId({ text: '秘密の自由入力値' }), 'other');
});

test('work log deltas accumulate counts and active time without storing input text', () => {
  const delta = createEmptyWorkLogDelta({ sessionCount: 1 });
  delta.totalActiveMs = 65000;
  addWorkActionDelta(delta, 'panel_edit', { count: 2, activeMs: 45000 });
  addWorkActionDelta(delta, 'panel_edit', { count: 1, activeMs: 5000 });
  const record = applyWorkLogDeltaToRecord({}, delta);

  assert.equal(record.totalActiveMs, 65000);
  assert.equal(record.sessionCount, 1);
  assert.deepEqual(record.actionStats.panel_edit, { label: 'コマ編集', count: 3, activeMs: 50000 });
  assert.equal(JSON.stringify(record).includes('秘密'), false);
});

test('work log records aggregate per account and sort by total active time', () => {
  const accounts = aggregateWorkLogRecords([
    { uid: 'a', email: 'a@example.com', displayName: 'A', totalActiveMs: 60000, sessionCount: 1, actionStats: { viewing: { count: 1, activeMs: 60000 } }, lastSeenAtMs: 100 },
    { uid: 'b', email: 'b@example.com', displayName: 'B', totalActiveMs: 240000, sessionCount: 2, actionStats: { navigation: { count: 3, activeMs: 120000 } }, lastSeenAtMs: 300 },
    { uid: 'a', email: 'a@example.com', displayName: 'A', totalActiveMs: 120000, sessionCount: 1, actionStats: { panel_edit: { count: 2, activeMs: 90000 } }, lastSeenAtMs: 200 }
  ]);

  assert.deepEqual(accounts.map((account) => account.uid), ['b', 'a']);
  assert.equal(accounts[1].totalActiveMs, 180000);
  assert.equal(accounts[1].sessionCount, 2);
  assert.equal(accounts[1].actions.panel_edit.count, 2);
});

test('work log date and duration formatting are deterministic', () => {
  assert.equal(getJstDateKey(new Date('2026-08-26T15:30:00.000Z')), '2026-08-27');
  assert.equal(formatWorkDuration(0), '0分');
  assert.equal(formatWorkDuration(65 * 60000), '1時間5分');
});
