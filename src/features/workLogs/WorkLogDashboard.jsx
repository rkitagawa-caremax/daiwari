import React, { useMemo, useState } from 'react';
import { Clock3, MousePointerClick, RefreshCw, Users, X } from 'lucide-react';

import { aggregateWorkLogRecords, formatWorkDuration, getJstDateKey } from '../../domain/workActivity';

const getCutoffDateKey = (rangeDays) => {
  if (rangeDays === 'all') return '';
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - Number(rangeDays) + 1);
  return getJstDateKey(cutoff);
};

const formatLastSeen = (milliseconds) => {
  if (!milliseconds) return '記録なし';
  return new Intl.DateTimeFormat('ja-JP', {
    timeZone: 'Asia/Tokyo',
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(new Date(milliseconds));
};

const WorkLogDashboard = React.memo(({
  isOpen,
  onClose,
  records,
  isLoading,
  errorMessage,
  onRefresh,
  isLocalMode
}) => {
  const [rangeDays, setRangeDays] = useState('30');
  const cutoffDateKey = getCutoffDateKey(rangeDays);
  const filteredRecords = useMemo(() => (
    (records || []).filter((record) => !cutoffDateKey || (record.dateKey || '') >= cutoffDateKey)
  ), [records, cutoffDateKey]);
  const accounts = useMemo(() => aggregateWorkLogRecords(filteredRecords), [filteredRecords]);
  const totalActiveMs = accounts.reduce((total, account) => total + account.totalActiveMs, 0);
  const totalActions = accounts.reduce((total, account) => (
    total + Object.values(account.actions).reduce((actionTotal, stats) => actionTotal + stats.count, 0)
  ), 0);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[145] flex items-center justify-center bg-slate-950/55 p-4 backdrop-blur-sm" onClick={onClose}>
      <div className="flex max-h-[92vh] w-full max-w-6xl flex-col overflow-hidden rounded-3xl border border-slate-200 bg-slate-50 shadow-2xl" onClick={(event) => event.stopPropagation()}>
        <header className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 bg-white px-6 py-4">
          <div>
            <h2 className="text-xl font-black text-slate-800">作業ログ ダッシュボード</h2>
            <p className="mt-0.5 text-xs text-slate-500">アカウント別のアクティブ作業時間と操作回数</p>
          </div>
          <div className="flex items-center gap-2">
            <select
              value={rangeDays}
              onChange={(event) => setRangeDays(event.target.value)}
              className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-bold text-slate-700"
              aria-label="ログ集計期間"
            >
              <option value="7">直近7日</option>
              <option value="30">直近30日</option>
              <option value="90">直近90日</option>
              <option value="all">全期間</option>
            </select>
            <button
              type="button"
              onClick={onRefresh}
              disabled={isLoading}
              className="rounded-xl border border-slate-200 bg-white p-2 text-slate-600 transition hover:bg-slate-100 disabled:opacity-40"
              title="ログを更新"
            >
              <RefreshCw size={18} className={isLoading ? 'animate-spin' : ''} />
            </button>
            <button type="button" onClick={onClose} className="rounded-xl p-2 text-slate-500 transition hover:bg-slate-100" aria-label="閉じる">
              <X size={20} />
            </button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-6">
          {isLocalMode && (
            <div className="mb-4 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-xs font-bold text-amber-800">
              ローカルモードの記録です。Googleアカウント別のクラウドログとは共有されません。
            </div>
          )}
          {errorMessage && (
            <div className="mb-4 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-xs font-bold text-rose-700">{errorMessage}</div>
          )}

          <div className="mb-6 grid gap-3 sm:grid-cols-3">
            {[
              { icon: Users, label: 'アカウント数', value: `${accounts.length}件`, color: 'text-indigo-600 bg-indigo-50' },
              { icon: Clock3, label: '総作業時間', value: formatWorkDuration(totalActiveMs), color: 'text-emerald-600 bg-emerald-50' },
              { icon: MousePointerClick, label: '総操作回数', value: `${totalActions.toLocaleString()}回`, color: 'text-sky-600 bg-sky-50' }
            ].map((summary) => (
              <div key={summary.label} className="flex items-center gap-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
                <div className={`rounded-xl p-3 ${summary.color}`}><summary.icon size={22} /></div>
                <div>
                  <p className="text-[11px] font-bold text-slate-500">{summary.label}</p>
                  <p className="text-xl font-black text-slate-800">{summary.value}</p>
                </div>
              </div>
            ))}
          </div>

          {isLoading && accounts.length === 0 ? (
            <div className="py-16 text-center text-sm font-bold text-slate-500">ログを読み込んでいます...</div>
          ) : accounts.length === 0 ? (
            <div className="rounded-2xl border border-dashed border-slate-300 bg-white py-16 text-center text-sm font-bold text-slate-500">
              選択期間の作業ログはまだありません。
            </div>
          ) : (
            <div className="space-y-4">
              {accounts.map((account) => {
                const actionRows = Object.entries(account.actions)
                  .map(([id, stats]) => ({ id, ...stats }))
                  .sort((left, right) => right.activeMs - left.activeMs || right.count - left.count);
                const accountActionCount = actionRows.reduce((total, stats) => total + stats.count, 0);
                return (
                  <section key={account.uid} className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
                    <div className="grid gap-3 border-b border-slate-100 px-5 py-4 sm:grid-cols-[minmax(180px,1fr)_150px_120px_150px] sm:items-center">
                      <div>
                        <h3 className="font-black text-slate-800">{account.displayName}</h3>
                        <p className="truncate text-[11px] text-slate-500">{account.email || account.uid}</p>
                      </div>
                      <div><p className="text-[10px] font-bold text-slate-400">総作業時間</p><p className="font-black text-emerald-700">{formatWorkDuration(account.totalActiveMs)}</p></div>
                      <div><p className="text-[10px] font-bold text-slate-400">操作回数</p><p className="font-black text-slate-700">{accountActionCount.toLocaleString()}回</p></div>
                      <div><p className="text-[10px] font-bold text-slate-400">最終記録</p><p className="text-xs font-bold text-slate-600">{formatLastSeen(account.lastSeenAtMs)}</p></div>
                    </div>
                    <div className="grid gap-2 p-4 sm:grid-cols-2 lg:grid-cols-3">
                      {actionRows.map((action) => (
                        <div key={action.id} className="rounded-xl bg-slate-50 px-3 py-2.5">
                          <p className="truncate text-xs font-black text-slate-700">{action.label}</p>
                          <div className="mt-1 flex justify-between text-[11px] font-bold text-slate-500">
                            <span>{action.count.toLocaleString()}回</span>
                            <span>{formatWorkDuration(action.activeMs)}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </section>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
});

export default WorkLogDashboard;
