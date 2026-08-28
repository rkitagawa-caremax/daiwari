import { useCallback, useEffect, useRef, useState } from 'react';

import {
  createAccountUndoStore,
  isSameUndoDomainList
} from '../features/undo/accountUndo';

export const useWorkspaceUndoState = ({ accountId }) => {
  const [sheets, setSheetsState] = useState([]);
  const [images, setImagesState] = useState([]);
  const [tempItems, setTempItemsState] = useState([]);
  const [excludedItems, setExcludedItemsState] = useState([]);
  const workspaceStateRef = useRef({
    sheets: [],
    images: [],
    tempItems: [],
    excludedItems: []
  });
  const undoStoreRef = useRef(null);
  const isUndoApplyingRef = useRef(false);
  const isUndoBusyRef = useRef(false);

  if (undoStoreRef.current == null) {
    undoStoreRef.current = createAccountUndoStore();
  }

  const replaceWorkspaceState = useCallback((domain, stateSetter, update, shouldTrack) => {
    const beforeItems = workspaceStateRef.current[domain] || [];
    const afterItems = typeof update === 'function' ? update(beforeItems) : update;
    const normalizedAfterItems = Array.isArray(afterItems) ? afterItems : [];
    if (isSameUndoDomainList(domain, beforeItems, normalizedAfterItems)) return beforeItems;

    workspaceStateRef.current = {
      ...workspaceStateRef.current,
      [domain]: normalizedAfterItems
    };
    if (shouldTrack && !isUndoApplyingRef.current) {
      undoStoreRef.current.record(domain, beforeItems, normalizedAfterItems);
    }
    stateSetter(normalizedAfterItems);
    return normalizedAfterItems;
  }, []);

  const setSheets = useCallback((update) => replaceWorkspaceState('sheets', setSheetsState, update, true), [replaceWorkspaceState]);
  const setImages = useCallback((update) => replaceWorkspaceState('images', setImagesState, update, true), [replaceWorkspaceState]);
  const setTempItems = useCallback((update) => replaceWorkspaceState('tempItems', setTempItemsState, update, true), [replaceWorkspaceState]);
  const setExcludedItems = useCallback((update) => replaceWorkspaceState('excludedItems', setExcludedItemsState, update, true), [replaceWorkspaceState]);
  const syncSheets = useCallback((update) => replaceWorkspaceState('sheets', setSheetsState, update, false), [replaceWorkspaceState]);
  const syncImages = useCallback((update) => replaceWorkspaceState('images', setImagesState, update, false), [replaceWorkspaceState]);
  const syncTempItems = useCallback((update) => replaceWorkspaceState('tempItems', setTempItemsState, update, false), [replaceWorkspaceState]);
  const syncExcludedItems = useCallback((update) => replaceWorkspaceState('excludedItems', setExcludedItemsState, update, false), [replaceWorkspaceState]);
  const flushPendingUndo = useCallback(() => undoStoreRef.current.flush(), []);
  const getLatestUndoEntry = useCallback((targetAccountId) => undoStoreRef.current.peekLatest(targetAccountId), []);
  const removeUndoEntry = useCallback((targetAccountId, entryId) => undoStoreRef.current.remove(targetAccountId, entryId), []);

  useEffect(() => {
    undoStoreRef.current.setAccount(accountId);
  }, [accountId]);

  useEffect(() => {
    const handleActionBoundary = (event) => {
      if (event.isPrimary === false || (event.button !== undefined && event.button !== 0)) return;
      if (undoStoreRef.current.hasPending()) undoStoreRef.current.flush();
    };
    document.addEventListener('pointerdown', handleActionBoundary, true);
    return () => document.removeEventListener('pointerdown', handleActionBoundary, true);
  }, []);

  useEffect(() => () => undoStoreRef.current.dispose(), []);

  return {
    excludedItems,
    flushPendingUndo,
    getLatestUndoEntry,
    images,
    isUndoApplyingRef,
    isUndoBusyRef,
    removeUndoEntry,
    setExcludedItems,
    setImages,
    setSheets,
    setTempItems,
    sheets,
    syncExcludedItems,
    syncImages,
    syncSheets,
    syncTempItems,
    tempItems,
    workspaceStateRef
  };
};
