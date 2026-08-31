import { useCallback, useEffect, useRef, useState } from 'react';

const LOCK_HOLD_MS = 2000;

export const useScreenLock = ({ isToggleDisabled = false } = {}) => {
  const [isLocked, setIsLocked] = useState(false);
  const lockHoldTimerRef = useRef(null);
  const lockHoldFiredRef = useRef(false);

  const cancelLockHold = useCallback(() => {
    if (!lockHoldTimerRef.current) return;
    clearTimeout(lockHoldTimerRef.current);
    lockHoldTimerRef.current = null;
  }, []);

  const startLockHold = useCallback(() => {
    if (isToggleDisabled) return;
    cancelLockHold();
    lockHoldFiredRef.current = false;
    lockHoldTimerRef.current = setTimeout(() => {
      lockHoldFiredRef.current = true;
      setIsLocked((current) => !current);
      lockHoldTimerRef.current = null;
    }, LOCK_HOLD_MS);
  }, [cancelLockHold, isToggleDisabled]);

  useEffect(() => cancelLockHold, [cancelLockHold]);

  return {
    isLocked,
    lockHoldFiredRef,
    startLockHold,
    cancelLockHold
  };
};
