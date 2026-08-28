export const UNDO_WORKSPACE_DOMAINS = [
  'sheets',
  'images',
  'tempItems',
  'excludedItems'
];

const getStableId = (item) => String(item?.id || '');

export const buildUndoDomainChanges = (
  beforeItems = [],
  afterItems = [],
  isSameItem = (left, right) => left === right
) => {
  const beforeById = new Map();
  const afterById = new Map();

  beforeItems.forEach((item, index) => {
    const id = getStableId(item);
    if (id) beforeById.set(id, { item, index });
  });
  afterItems.forEach((item, index) => {
    const id = getStableId(item);
    if (id) afterById.set(id, { item, index });
  });

  const changes = [];
  const ids = new Set([...beforeById.keys(), ...afterById.keys()]);
  ids.forEach((id) => {
    const beforeRecord = beforeById.get(id) || null;
    const afterRecord = afterById.get(id) || null;
    if (
      beforeRecord
      && afterRecord
      && isSameItem(beforeRecord.item, afterRecord.item)
    ) return;

    changes.push({
      id,
      before: beforeRecord?.item || null,
      after: afterRecord?.item || null,
      beforeIndex: beforeRecord?.index ?? -1,
      afterIndex: afterRecord?.index ?? -1
    });
  });

  return changes;
};

export const mergeUndoDomainChanges = (
  existingChanges = [],
  incomingChanges = [],
  isSameItem = (left, right) => left === right
) => {
  const merged = new Map(existingChanges.map((change) => [change.id, change]));

  incomingChanges.forEach((change) => {
    const existing = merged.get(change.id);
    if (!existing) {
      merged.set(change.id, change);
      return;
    }

    const next = {
      ...existing,
      after: change.after,
      afterIndex: change.afterIndex
    };
    const returnedToOriginalValue = next.before === null
      ? next.after === null
      : next.after !== null && isSameItem(next.before, next.after);
    const returnedToOriginalIndex = next.beforeIndex === next.afterIndex;

    if (returnedToOriginalValue && returnedToOriginalIndex) {
      merged.delete(change.id);
    } else {
      merged.set(change.id, next);
    }
  });

  return Array.from(merged.values());
};

export const applyUndoDomainChanges = (currentItems = [], changes = []) => {
  const changesById = new Map(changes.map((change) => [change.id, change]));
  const restored = currentItems
    .filter((item) => {
      const change = changesById.get(getStableId(item));
      return !change || change.before !== null;
    })
    .map((item) => {
      const change = changesById.get(getStableId(item));
      return change?.before || item;
    });

  changes
    .filter((change) => change.before !== null && !restored.some((item) => getStableId(item) === change.id))
    .sort((left, right) => left.beforeIndex - right.beforeIndex)
    .forEach((change) => {
      const targetIndex = Math.max(0, Math.min(change.beforeIndex, restored.length));
      restored.splice(targetIndex, 0, change.before);
    });

  const originalOrder = new Map(changes
    .filter((change) => change.before !== null)
    .map((change) => [change.id, change.beforeIndex]));

  if (originalOrder.size > 0) {
    const stablePositions = new Map(restored.map((item, index) => [getStableId(item), index]));
    restored.sort((left, right) => {
      const leftId = getStableId(left);
      const rightId = getStableId(right);
      const leftPosition = originalOrder.has(leftId) ? originalOrder.get(leftId) : stablePositions.get(leftId);
      const rightPosition = originalOrder.has(rightId) ? originalOrder.get(rightId) : stablePositions.get(rightId);
      return leftPosition - rightPosition;
    });
  }

  return restored;
};

export const hasUndoEntryChanges = (entry) => (
  UNDO_WORKSPACE_DOMAINS.some((domain) => (entry?.changes?.[domain]?.length || 0) > 0)
);
