export const WORK_ACTIONS = Object.freeze({
  viewing: { label: '閲覧・確認' },
  navigation: { label: '画面・ページ移動' },
  panel_edit: { label: 'コマ編集' },
  page_operation: { label: 'ページ操作' },
  image_operation: { label: '画像・素材操作' },
  label_operation: { label: 'ラベル操作' },
  data_io: { label: 'CSV入出力' },
  pdf_export: { label: 'PDF出力' },
  settings: { label: '設定・管理' },
  other: { label: 'その他の操作' }
});

const ACTION_IDS = new Set(Object.keys(WORK_ACTIONS));

export const getJstDateKey = (date = new Date()) => {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Tokyo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
};

export const normalizeWorkActionId = (value) => (
  ACTION_IDS.has(value) ? value : 'other'
);

export const resolveWorkActionId = ({ explicitAction, text = '', eventType = '', isPanel = false } = {}) => {
  if (ACTION_IDS.has(explicitAction)) return explicitAction;
  if (isPanel) return 'panel_edit';
  if (eventType === 'drop' || eventType === 'dragstart') return 'image_operation';

  const descriptor = String(text || '').toLowerCase();
  if (/pdf/.test(descriptor)) return 'pdf_export';
  if (/csv|取込|取り込|インポート|エクスポート|出力/.test(descriptor)) return 'data_io';
  if (/ラベル/.test(descriptor)) return 'label_operation';
  if (/画像|素材|ライブラリ|仮置|除外|アップロード|ダミー/.test(descriptor)) return 'image_operation';
  if (/ページ追加|ページ削除|入れ替え|選択モード|全選択|画像解除/.test(descriptor)) return 'page_operation';
  if (/設定|ログ|管理|ロック|バーを/.test(descriptor)) return 'settings';
  if (/詳細|全体|ジャンル|次のページ|前のページ|空き|検索/.test(descriptor)) return 'navigation';
  return 'other';
};

export const createEmptyWorkLogDelta = ({ sessionCount = 0 } = {}) => ({
  totalActiveMs: 0,
  sessionCount,
  actions: {}
});

export const addWorkActionDelta = (delta, actionId, { count = 0, activeMs = 0 } = {}) => {
  const normalizedId = normalizeWorkActionId(actionId);
  const previous = delta.actions[normalizedId] || { count: 0, activeMs: 0 };
  delta.actions[normalizedId] = {
    count: previous.count + Math.max(0, Number(count) || 0),
    activeMs: previous.activeMs + Math.max(0, Number(activeMs) || 0)
  };
  return delta;
};

export const mergeWorkLogDelta = (target, incoming) => {
  target.totalActiveMs += Math.max(0, Number(incoming?.totalActiveMs) || 0);
  target.sessionCount += Math.max(0, Number(incoming?.sessionCount) || 0);
  Object.entries(incoming?.actions || {}).forEach(([actionId, stats]) => {
    addWorkActionDelta(target, actionId, stats);
  });
  return target;
};

export const applyWorkLogDeltaToRecord = (record = {}, delta = {}) => {
  const actionStats = { ...(record.actionStats || {}) };
  Object.entries(delta.actions || {}).forEach(([actionId, stats]) => {
    const previous = actionStats[actionId] || {};
    actionStats[actionId] = {
      label: WORK_ACTIONS[normalizeWorkActionId(actionId)].label,
      count: (Number(previous.count) || 0) + (Number(stats.count) || 0),
      activeMs: (Number(previous.activeMs) || 0) + (Number(stats.activeMs) || 0)
    };
  });
  return {
    ...record,
    totalActiveMs: (Number(record.totalActiveMs) || 0) + (Number(delta.totalActiveMs) || 0),
    sessionCount: (Number(record.sessionCount) || 0) + (Number(delta.sessionCount) || 0),
    actionStats
  };
};

const timestampToMillis = (value) => {
  if (typeof value?.toMillis === 'function') return value.toMillis();
  if (typeof value?.seconds === 'number') return value.seconds * 1000;
  return Number(value) || 0;
};

export const aggregateWorkLogRecords = (records = []) => {
  const accounts = new Map();
  (Array.isArray(records) ? records : []).forEach((record) => {
    if (!record?.uid) return;
    const account = accounts.get(record.uid) || {
      uid: record.uid,
      email: record.email || '',
      displayName: record.displayName || record.email || '不明なアカウント',
      totalActiveMs: 0,
      sessionCount: 0,
      lastSeenAtMs: 0,
      actions: {}
    };
    account.email = record.email || account.email;
    account.displayName = record.displayName || account.displayName;
    account.totalActiveMs += Math.max(0, Number(record.totalActiveMs) || 0);
    account.sessionCount += Math.max(0, Number(record.sessionCount) || 0);
    account.lastSeenAtMs = Math.max(account.lastSeenAtMs, timestampToMillis(record.lastSeenAt || record.lastSeenAtMs));
    Object.entries(record.actionStats || {}).forEach(([actionId, stats]) => {
      const normalizedId = normalizeWorkActionId(actionId);
      const previous = account.actions[normalizedId] || {
        label: WORK_ACTIONS[normalizedId].label,
        count: 0,
        activeMs: 0
      };
      account.actions[normalizedId] = {
        ...previous,
        count: previous.count + Math.max(0, Number(stats?.count) || 0),
        activeMs: previous.activeMs + Math.max(0, Number(stats?.activeMs) || 0)
      };
    });
    accounts.set(record.uid, account);
  });
  return Array.from(accounts.values()).sort((left, right) => right.totalActiveMs - left.totalActiveMs);
};

export const formatWorkDuration = (milliseconds = 0) => {
  const totalMinutes = Math.floor(Math.max(0, Number(milliseconds) || 0) / 60000);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (hours > 0) return `${hours}時間${minutes}分`;
  return `${minutes}分`;
};
