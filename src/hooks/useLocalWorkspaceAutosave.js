import { useEffect, useRef } from 'react';

import { idbHelper } from '../idbHelper.js';

export const LOCAL_WORKSPACE_AUTOSAVE_DELAY_MS = 500;

export const createLocalWorkspaceSavedSnapshot = () => ({
  sheets: null,
  images: null,
  tempItems: null,
  excludedItems: null,
  salesData: null
});

const reportAutosaveError = (error) => {
  console.error('Auto-save failed:', error);
};

const clearPendingAutosave = (timeoutRef, cancelTimer) => {
  if (timeoutRef.current) {
    cancelTimer(timeoutRef.current);
  }
};

export const scheduleLocalWorkspaceAutosave = ({
  isLocalStorageMode,
  isDataLoaded,
  snapshot,
  savedSnapshotRef,
  timeoutRef,
  storage = idbHelper,
  scheduleTimer = globalThis.setTimeout,
  cancelTimer = globalThis.clearTimeout,
  delay = LOCAL_WORKSPACE_AUTOSAVE_DELAY_MS,
  onError = reportAutosaveError
}) => {
  const cancelPending = () => clearPendingAutosave(timeoutRef, cancelTimer);

  if (!(isLocalStorageMode && isDataLoaded)) {
    return cancelPending;
  }

  cancelPending();

  const shouldSaveSheets = savedSnapshotRef.current.sheets !== snapshot.sheets;
  const shouldSaveImages = savedSnapshotRef.current.images !== snapshot.images;
  const shouldSaveTempItems = savedSnapshotRef.current.tempItems !== snapshot.tempItems;
  const shouldSaveExcludedItems = savedSnapshotRef.current.excludedItems !== snapshot.excludedItems;
  const shouldSaveSalesData = savedSnapshotRef.current.salesData !== snapshot.salesData;

  if (!shouldSaveSheets && !shouldSaveImages && !shouldSaveTempItems && !shouldSaveExcludedItems && !shouldSaveSalesData) {
    return cancelPending;
  }

  timeoutRef.current = scheduleTimer(async () => {
    try {
      const tasks = [];
      if (shouldSaveSheets) {
        tasks.push(storage.setItem('sheets', snapshot.sheets).then(() => {
          savedSnapshotRef.current.sheets = snapshot.sheets;
        }));
      }
      if (shouldSaveImages) {
        tasks.push(storage.setItem('images', snapshot.images).then(() => {
          savedSnapshotRef.current.images = snapshot.images;
        }));
      }
      if (shouldSaveTempItems) {
        tasks.push(storage.setItem('tempItems', snapshot.tempItems).then(() => {
          savedSnapshotRef.current.tempItems = snapshot.tempItems;
        }));
      }
      if (shouldSaveExcludedItems) {
        tasks.push(storage.setItem('excludedItems', snapshot.excludedItems).then(() => {
          savedSnapshotRef.current.excludedItems = snapshot.excludedItems;
        }));
      }
      if (shouldSaveSalesData && snapshot.salesData) {
        tasks.push(storage.setItem('salesData', snapshot.salesData).then(() => {
          savedSnapshotRef.current.salesData = snapshot.salesData;
        }));
      }
      await Promise.all(tasks);
    } catch (error) {
      onError(error);
    }
  }, delay);

  return cancelPending;
};

export const useLocalWorkspaceAutosave = ({
  isLocalStorageMode,
  isDataLoaded,
  sheets,
  images,
  tempItems,
  excludedItems,
  salesData,
  storage = idbHelper,
  scheduleTimer = globalThis.setTimeout,
  cancelTimer = globalThis.clearTimeout,
  delay = LOCAL_WORKSPACE_AUTOSAVE_DELAY_MS,
  onError = reportAutosaveError
}) => {
  const timeoutRef = useRef(null);
  const savedSnapshotRef = useRef(createLocalWorkspaceSavedSnapshot());

  useEffect(() => scheduleLocalWorkspaceAutosave({
    isLocalStorageMode,
    isDataLoaded,
    snapshot: { sheets, images, tempItems, excludedItems, salesData },
    savedSnapshotRef,
    timeoutRef,
    storage,
    scheduleTimer,
    cancelTimer,
    delay,
    onError
  }), [
    cancelTimer,
    delay,
    excludedItems,
    images,
    isDataLoaded,
    isLocalStorageMode,
    onError,
    salesData,
    scheduleTimer,
    sheets,
    storage,
    tempItems
  ]);
};
