import { deleteField, doc, serverTimestamp } from 'firebase/firestore';

import {
  buildDefaultPanels,
  getPanelFreeLabels,
  getPanelsFromDocData,
  toPanelsMap
} from '../../domain/panels';
import { isSameStockImageList, normalizeStockImages } from '../../domain/images';
import {
  UNDO_WORKSPACE_DOMAINS,
  applyUndoDomainChanges,
  buildUndoDomainChanges,
  hasUndoEntryChanges,
  mergeUndoDomainChanges
} from '../../domain/undoHistory';
import {
  isSameSheetList,
  isSameTransferItemList
} from '../../domain/workspaceComparators';

export { UNDO_WORKSPACE_DOMAINS };

const isSameUndoTransferItem = (left, right) => isSameTransferItemList(
  [{ ...(left || {}), createdAt: null }],
  [{ ...(right || {}), createdAt: null }]
);

export const isSameUndoDomainItem = (domain, left, right) => {
  if (left === null || right === null) return left === right;
  if (domain === 'sheets') return isSameSheetList([left], [right]);
  if (domain === 'images') return isSameStockImageList([left], [right]);
  return isSameUndoTransferItem(left, right);
};

export const isSameUndoDomainList = (domain, leftItems, rightItems) => {
  if (domain === 'sheets') return isSameSheetList(leftItems, rightItems);
  if (domain === 'images') return isSameStockImageList(leftItems, rightItems);
  return isSameTransferItemList(leftItems, rightItems);
};

export const createAccountUndoStore = ({
  groupDelayMs = 2000,
  maxEntriesPerAccount = 30
} = {}) => {
  const historyByAccount = new Map();
  let activeAccountId = null;
  let pendingEntry = null;
  let flushTimer = null;

  const clearTimer = () => {
    if (flushTimer) clearTimeout(flushTimer);
    flushTimer = null;
  };

  const flush = () => {
    clearTimer();
    const entry = pendingEntry;
    pendingEntry = null;
    if (!entry?.accountId || !hasUndoEntryChanges(entry)) return null;

    const history = historyByAccount.get(entry.accountId) || [];
    history.push(entry);
    if (history.length > maxEntriesPerAccount) {
      history.splice(0, history.length - maxEntriesPerAccount);
    }
    historyByAccount.set(entry.accountId, history);
    return entry;
  };

  const setAccount = (accountId) => {
    if (activeAccountId === accountId) return;
    flush();
    activeAccountId = accountId || null;
  };

  const record = (domain, beforeItems, afterItems) => {
    if (!activeAccountId || !UNDO_WORKSPACE_DOMAINS.includes(domain)) return;
    const incomingChanges = buildUndoDomainChanges(
      beforeItems,
      afterItems,
      (left, right) => isSameUndoDomainItem(domain, left, right)
    );
    if (incomingChanges.length === 0) return;

    if (pendingEntry && pendingEntry.accountId !== activeAccountId) flush();
    if (!pendingEntry) {
      pendingEntry = {
        id: `${activeAccountId}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        accountId: activeAccountId,
        createdAt: Date.now(),
        changes: Object.fromEntries(UNDO_WORKSPACE_DOMAINS.map((key) => [key, []]))
      };
    }

    pendingEntry.changes[domain] = mergeUndoDomainChanges(
      pendingEntry.changes[domain],
      incomingChanges,
      (left, right) => isSameUndoDomainItem(domain, left, right)
    );
    clearTimer();
    flushTimer = setTimeout(flush, groupDelayMs);
  };

  const peekLatest = (accountId = activeAccountId) => {
    flush();
    const history = historyByAccount.get(accountId) || [];
    return history[history.length - 1] || null;
  };

  const remove = (accountId, entryId) => {
    const history = historyByAccount.get(accountId) || [];
    historyByAccount.set(accountId, history.filter((entry) => entry.id !== entryId));
  };

  const dispose = () => {
    clearTimer();
    pendingEntry = null;
  };

  return {
    dispose,
    flush,
    hasPending: () => !!pendingEntry,
    peekLatest,
    record,
    remove,
    setAccount
  };
};

export const undoEntryHasClientConflict = (entry, workspaceState) => (
  UNDO_WORKSPACE_DOMAINS.some((domain) => {
    const currentItems = workspaceState[domain] || [];
    return (entry?.changes?.[domain] || []).some((change) => {
      const currentItem = currentItems.find((item) => String(item?.id || '') === change.id) || null;
      return !isSameUndoDomainItem(domain, currentItem, change.after);
    });
  })
);

export const countUndoEntryChanges = (entry) => UNDO_WORKSPACE_DOMAINS.reduce(
  (count, domain) => count + (entry?.changes?.[domain]?.length || 0),
  0
);

export const applyUndoEntryToWorkspace = (entry, workspaceState) => Object.fromEntries(
  UNDO_WORKSPACE_DOMAINS.map((domain) => [
    domain,
    applyUndoDomainChanges(workspaceState[domain] || [], entry?.changes?.[domain] || [])
  ])
);

const getSnapshotItem = (domain, snapshot) => {
  if (!snapshot?.exists()) return null;
  const data = snapshot.data() || {};
  if (domain === 'sheets') {
    return {
      ...data,
      id: snapshot.id,
      _hasLegacyPanels: Array.isArray(data.panels),
      panels: getPanelsFromDocData(data)
    };
  }
  if (domain === 'images') {
    return normalizeStockImages([{ ...data, id: snapshot.id }])[0] || null;
  }
  return { ...data, id: snapshot.id };
};

const getFirestoreData = (domain, item, options = {}) => {
  const createdAt = item?.createdAt || serverTimestamp();
  if (domain === 'sheets') {
    return {
      genre: item?.genre || 'none',
      order: item?.order ?? 0,
      panelsMap: toPanelsMap(item?.panels || buildDefaultPanels()),
      createdAt
    };
  }
  if (domain === 'images') {
    return {
      name: item?.name || item?.originalName || 'image.png',
      data: item?.data || item?.image || null,
      code: item?.code || null,
      freeLabels: getPanelFreeLabels(item),
      freeText: null,
      createdAt
    };
  }

  return {
    image: item?.image || item?.data || null,
    imageId: item?.imageId || null,
    label: item?.label || null,
    code: item?.code || null,
    text: item?.text || '',
    isText: !!item?.isText,
    freeLabels: getPanelFreeLabels(item),
    freeText: null,
    originalName: item?.originalName || item?.name || '退避アイテム',
    ...(domain === 'tempItems' && options.ownerUid ? { ownerUid: options.ownerUid } : {}),
    createdAt
  };
};

export const restoreCloudUndoEntry = async ({
  accountId,
  collections,
  entry,
  ownerUid,
  runCloudTransaction,
  useLegacyTempShelf
}) => {
  const targets = UNDO_WORKSPACE_DOMAINS.flatMap((domain) => (
    (entry?.changes?.[domain] || []).map((change) => ({
      domain,
      change,
      collectionRef: collections[domain]
    }))
  ));
  if (targets.some((target) => !target.collectionRef)) {
    throw Object.assign(new Error('Undo collection is not ready.'), { code: 'undo-not-ready' });
  }

  await runCloudTransaction(async (transaction) => {
    const snapshots = await Promise.all(targets.map((target) => (
      transaction.get(doc(target.collectionRef, target.change.id))
    )));

    snapshots.forEach((snapshot, index) => {
      const target = targets[index];
      const currentItem = getSnapshotItem(target.domain, snapshot);
      if (!isSameUndoDomainItem(target.domain, currentItem, target.change.after)) {
        throw Object.assign(new Error('Undo target changed remotely.'), { code: 'undo-conflict' });
      }
    });

    targets.forEach((target, index) => {
      const snapshot = snapshots[index];
      const targetRef = doc(target.collectionRef, target.change.id);
      if (target.change.before === null) {
        if (snapshot.exists()) transaction.delete(targetRef);
        return;
      }

      const payload = getFirestoreData(target.domain, target.change.before, {
        ownerUid: target.domain === 'tempItems' && useLegacyTempShelf ? ownerUid : null
      });
      if (!snapshot.exists()) {
        transaction.set(targetRef, payload);
      } else if (target.domain === 'sheets') {
        transaction.update(targetRef, { ...payload, panels: deleteField() });
      } else {
        transaction.set(targetRef, payload, { merge: true });
      }
    });
  }, { key: `undo:${accountId}` });
};
