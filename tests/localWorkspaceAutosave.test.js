import test from 'node:test';
import assert from 'node:assert/strict';

import {
  LOCAL_WORKSPACE_AUTOSAVE_DELAY_MS,
  createLocalWorkspaceSavedSnapshot,
  scheduleLocalWorkspaceAutosave
} from '../src/hooks/useLocalWorkspaceAutosave.js';

const createHarness = (overrides = {}) => {
  const scheduled = [];
  const cancelled = [];
  const writes = [];
  const errors = [];
  const savedSnapshotRef = { current: createLocalWorkspaceSavedSnapshot() };
  const timeoutRef = { current: null };
  const storage = {
    setItem: async (key, value) => {
      writes.push([key, value]);
    }
  };
  const scheduleTimer = (callback, delay) => {
    const id = scheduled.length + 1;
    scheduled.push({ id, callback, delay });
    return id;
  };
  const cancelTimer = (id) => cancelled.push(id);
  const snapshot = {
    sheets: [],
    images: [],
    tempItems: [],
    excludedItems: [],
    salesData: null
  };

  return {
    scheduled,
    cancelled,
    writes,
    errors,
    savedSnapshotRef,
    timeoutRef,
    storage,
    scheduleTimer,
    cancelTimer,
    snapshot,
    schedule: (options = {}) => scheduleLocalWorkspaceAutosave({
      isLocalStorageMode: true,
      isDataLoaded: true,
      snapshot,
      savedSnapshotRef,
      timeoutRef,
      storage,
      scheduleTimer,
      cancelTimer,
      onError: (error) => errors.push(error),
      ...overrides,
      ...options
    })
  };
};

test('local workspace autosave requires both local mode and loaded data', () => {
  const localDisabled = createHarness();
  localDisabled.timeoutRef.current = 7;
  const cleanupLocalDisabled = localDisabled.schedule({ isLocalStorageMode: false });
  assert.equal(localDisabled.scheduled.length, 0);
  cleanupLocalDisabled();
  assert.deepEqual(localDisabled.cancelled, [7]);

  const dataNotLoaded = createHarness();
  dataNotLoaded.schedule({ isDataLoaded: false });
  assert.equal(dataNotLoaded.scheduled.length, 0);
});

test('local workspace autosave waits 500ms and writes only changed references', async () => {
  const harness = createHarness();
  Object.assign(harness.savedSnapshotRef.current, harness.snapshot);
  harness.snapshot.images = [{ id: 'image-1' }];

  const cleanup = harness.schedule();

  assert.equal(harness.scheduled.length, 1);
  assert.equal(harness.scheduled[0].delay, LOCAL_WORKSPACE_AUTOSAVE_DELAY_MS);
  assert.deepEqual(harness.writes, []);

  await harness.scheduled[0].callback();

  assert.deepEqual(harness.writes, [['images', harness.snapshot.images]]);
  assert.equal(harness.savedSnapshotRef.current.images, harness.snapshot.images);
  cleanup();
  assert.deepEqual(harness.cancelled, [harness.scheduled[0].id]);
});

test('local workspace autosave keeps falsy sales data unsaved while saving other changes', async () => {
  const harness = createHarness();
  Object.assign(harness.savedSnapshotRef.current, harness.snapshot, { salesData: { before: true } });
  harness.snapshot.sheets = [{ id: 'sheet-1' }];
  harness.snapshot.salesData = null;

  harness.schedule();
  await harness.scheduled[0].callback();

  assert.deepEqual(harness.writes, [['sheets', harness.snapshot.sheets]]);
  assert.deepEqual(harness.savedSnapshotRef.current.salesData, { before: true });
});

test('local workspace autosave skips scheduling when every reference is already saved', () => {
  const harness = createHarness();
  Object.assign(harness.savedSnapshotRef.current, harness.snapshot);

  harness.schedule();

  assert.equal(harness.scheduled.length, 0);
  assert.deepEqual(harness.writes, []);
});

test('local workspace autosave reports storage failures through the injected handler', async () => {
  const expectedError = new Error('storage unavailable');
  const harness = createHarness({
    storage: {
      setItem: async () => {
        throw expectedError;
      }
    }
  });

  harness.schedule();
  await harness.scheduled[0].callback();

  assert.deepEqual(harness.errors, [expectedError]);
});
