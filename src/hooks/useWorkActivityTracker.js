import { useCallback, useEffect, useRef } from 'react';

import {
  WORK_ACTIONS,
  addWorkActionDelta,
  createEmptyWorkLogDelta,
  getJstDateKey,
  mergeWorkLogDelta,
  resolveWorkActionId
} from '../domain/workActivity';

const HEARTBEAT_MS = 30000;
const TICK_MS = 1000;
const IDLE_TIMEOUT_MS = 5 * 60 * 1000;
const ACTION_COUNT_DEBOUNCE_MS = 1500;

const getActionDescriptor = (target) => {
  if (!(target instanceof Element)) return { explicitAction: '', text: '', isPanel: false };
  const actionElement = target.closest('[data-work-action]');
  const interactiveElement = target.closest('button, input, textarea, select, [role="button"], [title], [aria-label]');
  return {
    explicitAction: actionElement?.dataset.workAction || '',
    text: [
      interactiveElement?.getAttribute('aria-label'),
      interactiveElement?.getAttribute('title'),
      interactiveElement?.getAttribute('placeholder'),
      interactiveElement?.textContent
    ].filter(Boolean).join(' ').slice(0, 160),
    isPanel: !!target.closest('[id^="panel-"]')
  };
};

export const useWorkActivityTracker = ({ enabled, user, onFlush }) => {
  const onFlushRef = useRef(onFlush);
  const pendingRef = useRef(createEmptyWorkLogDelta({ sessionCount: 1 }));
  const currentActionRef = useRef('viewing');
  const lastTickRef = useRef(Date.now());
  const lastInteractionRef = useRef(Date.now());
  const lastCountedActionRef = useRef({ id: '', at: 0 });
  const isFlushingRef = useRef(false);
  const mountedRef = useRef(false);

  useEffect(() => {
    onFlushRef.current = onFlush;
  }, [onFlush]);

  const collectActiveTime = useCallback(() => {
    const now = Date.now();
    const elapsed = Math.min(Math.max(0, now - lastTickRef.current), TICK_MS * 3);
    lastTickRef.current = now;
    const isActive = document.visibilityState === 'visible'
      && document.hasFocus()
      && now - lastInteractionRef.current <= IDLE_TIMEOUT_MS;
    if (!isActive || elapsed === 0) return;
    pendingRef.current.totalActiveMs += elapsed;
    addWorkActionDelta(pendingRef.current, currentActionRef.current, { activeMs: elapsed });
  }, []);

  const flushNow = useCallback(async () => {
    if (!enabled || !user?.uid || isFlushingRef.current) return;
    collectActiveTime();
    const current = pendingRef.current;
    const hasActions = Object.values(current.actions).some((stats) => stats.count > 0 || stats.activeMs > 0);
    if (current.totalActiveMs <= 0 && current.sessionCount <= 0 && !hasActions) return;

    pendingRef.current = createEmptyWorkLogDelta();
    isFlushingRef.current = true;
    try {
      await onFlushRef.current?.({
        ...current,
        dateKey: getJstDateKey(),
        user: {
          uid: user.uid,
          email: user.email || '',
          displayName: user.displayName || user.email || '不明なアカウント'
        }
      });
    } catch (error) {
      mergeWorkLogDelta(pendingRef.current, current);
      console.warn('Work activity log flush failed:', error);
    } finally {
      isFlushingRef.current = false;
    }
  }, [collectActiveTime, enabled, user?.displayName, user?.email, user?.uid]);

  useEffect(() => {
    if (!enabled || !user?.uid) return undefined;
    if (!mountedRef.current) {
      mountedRef.current = true;
      pendingRef.current = createEmptyWorkLogDelta({ sessionCount: 1 });
    }
    lastTickRef.current = Date.now();
    lastInteractionRef.current = Date.now();

    const recordInteraction = (event) => {
      collectActiveTime();
      const descriptor = getActionDescriptor(event.target);
      const actionId = resolveWorkActionId({ ...descriptor, eventType: event.type });
      const now = Date.now();
      currentActionRef.current = actionId;
      lastInteractionRef.current = now;
      const previous = lastCountedActionRef.current;
      if (previous.id !== actionId || now - previous.at >= ACTION_COUNT_DEBOUNCE_MS) {
        addWorkActionDelta(pendingRef.current, actionId, { count: 1 });
        lastCountedActionRef.current = { id: actionId, at: now };
      }
    };
    const handleVisibility = () => {
      collectActiveTime();
      lastTickRef.current = Date.now();
      if (document.visibilityState !== 'visible') void flushNow();
    };
    const handleFocus = () => {
      lastTickRef.current = Date.now();
      lastInteractionRef.current = Date.now();
    };
    const handleBlur = () => {
      collectActiveTime();
      void flushNow();
    };

    const tickTimer = window.setInterval(collectActiveTime, TICK_MS);
    const heartbeatTimer = window.setInterval(() => { void flushNow(); }, HEARTBEAT_MS);
    ['pointerdown', 'input', 'change', 'drop', 'dragstart'].forEach((eventName) => {
      document.addEventListener(eventName, recordInteraction, true);
    });
    document.addEventListener('visibilitychange', handleVisibility);
    window.addEventListener('focus', handleFocus);
    window.addEventListener('blur', handleBlur);

    return () => {
      window.clearInterval(tickTimer);
      window.clearInterval(heartbeatTimer);
      ['pointerdown', 'input', 'change', 'drop', 'dragstart'].forEach((eventName) => {
        document.removeEventListener(eventName, recordInteraction, true);
      });
      document.removeEventListener('visibilitychange', handleVisibility);
      window.removeEventListener('focus', handleFocus);
      window.removeEventListener('blur', handleBlur);
      collectActiveTime();
      void flushNow();
    };
  }, [collectActiveTime, enabled, flushNow, user?.uid]);

  return { flushWorkActivityNow: flushNow, workActionDefinitions: WORK_ACTIONS };
};
