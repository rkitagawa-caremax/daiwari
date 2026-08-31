import { useCallback, useEffect, useRef, useState } from 'react';

import {
  DAIWARI_DROPZONE_ATTR,
  POINTER_DRAG_THRESHOLD_PX,
  clearActiveNativeDragPayload,
  getActiveNativeDragPayload,
  getDragPayload,
  isDropEventHandled,
  markDropEventHandled,
  normalizeDragPayload
} from '../lib/dragPayload';
import { dispatchWorkspaceDrop } from '../lib/workspaceDropZones';

export const useWorkspacePointerDrag = ({
  onDropToPanel,
  onDropToTemp,
  onDropToStock,
  onDropToExcluded,
  suppressNextClickRef
}) => {
  const [pointerDragPreview, setPointerDragPreview] = useState(null);
  const pointerDragOverlayRef = useRef(null);
  const pointerDragSessionRef = useRef(null);

  const dispatchDropToZone = useCallback((zoneId, dragPayload = {}) => (
    dispatchWorkspaceDrop({
      zoneId,
      dragPayload,
      onDropToPanel,
      onDropToTemp,
      onDropToStock,
      onDropToExcluded
    })
  ), [onDropToExcluded, onDropToPanel, onDropToStock, onDropToTemp]);

  const dropAtPoint = useCallback((clientX, clientY, dragPayload = {}) => {
    if (typeof document === 'undefined') return false;
    const target = document.elementFromPoint(clientX, clientY);
    const dropZoneElement = target?.closest?.(`[${DAIWARI_DROPZONE_ATTR}]`);
    const zoneId = dropZoneElement?.getAttribute?.(DAIWARI_DROPZONE_ATTR) || '';
    return dispatchDropToZone(zoneId, dragPayload);
  }, [dispatchDropToZone]);

  useEffect(() => {
    if (typeof document === 'undefined') return undefined;

    const handleDocumentDragOver = (event) => {
      if (!getActiveNativeDragPayload()) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
    };

    const handleDocumentDrop = (event) => {
      if (isDropEventHandled(event)) return;
      const dragPayload = getDragPayload(event.dataTransfer);
      if (!dragPayload) {
        clearActiveNativeDragPayload();
        return;
      }

      const targetElement = event.target instanceof Element
        ? event.target
        : event.target?.parentElement;
      const dropZoneElement = targetElement?.closest?.(`[${DAIWARI_DROPZONE_ATTR}]`);
      const zoneId = dropZoneElement?.getAttribute?.(DAIWARI_DROPZONE_ATTR) || '';
      if (!zoneId) {
        clearActiveNativeDragPayload();
        return;
      }

      const handled = dispatchDropToZone(zoneId, dragPayload);
      if (handled) {
        markDropEventHandled(event);
        event.preventDefault();
        event.stopPropagation();
      }
      clearActiveNativeDragPayload();
    };

    document.addEventListener('dragover', handleDocumentDragOver);
    document.addEventListener('drop', handleDocumentDrop);
    return () => {
      document.removeEventListener('dragover', handleDocumentDragOver);
      document.removeEventListener('drop', handleDocumentDrop);
    };
  }, [dispatchDropToZone]);

  const positionOverlay = useCallback((clientX, clientY) => {
    const overlay = pointerDragOverlayRef.current;
    if (!overlay) return;
    overlay.style.transform = `translate3d(${clientX + 18}px, ${clientY + 18}px, 0)`;
  }, []);

  const clearPointerDragSession = useCallback(() => {
    const session = pointerDragSessionRef.current;
    if (!session || typeof document === 'undefined') {
      pointerDragSessionRef.current = null;
      setPointerDragPreview(null);
      return;
    }

    document.removeEventListener('pointermove', session.handleMove);
    document.removeEventListener('pointerup', session.handleUp);
    document.removeEventListener('pointercancel', session.handleCancel);
    pointerDragSessionRef.current = null;
    setPointerDragPreview(null);
  }, []);

  const startPointerDrag = useCallback((event, config = {}) => {
    if (!event || event.pointerType === 'mouse' || event.isPrimary === false || !config.payload) return;
    if (event.button !== undefined && event.button !== 0) return;

    clearPointerDragSession();
    const session = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      currentX: event.clientX,
      currentY: event.clientY,
      active: false,
      payload: normalizeDragPayload(config.payload),
      preview: config.preview || {}
    };

    const finishPointerDrag = (pointerEvent, shouldDrop) => {
      if (!pointerEvent || pointerEvent.pointerId !== session.pointerId) return;
      const wasActive = session.active;
      clearPointerDragSession();
      let handled = false;
      if (wasActive) {
        suppressNextClickRef.current = true;
        if (shouldDrop) {
          handled = dropAtPoint(pointerEvent.clientX, pointerEvent.clientY, session.payload);
        }
      }
      config.onFinish?.({ active: wasActive, handled });
    };

    session.handleMove = (moveEvent) => {
      if (!moveEvent || moveEvent.pointerId !== session.pointerId) return;
      session.currentX = moveEvent.clientX;
      session.currentY = moveEvent.clientY;

      if (!session.active) {
        const distance = Math.hypot(
          session.currentX - session.startX,
          session.currentY - session.startY
        );
        if (distance < POINTER_DRAG_THRESHOLD_PX) return;
        session.active = true;
        setPointerDragPreview(session.preview);
        requestAnimationFrame(() => positionOverlay(session.currentX, session.currentY));
      }

      moveEvent.preventDefault();
      positionOverlay(session.currentX, session.currentY);
    };
    session.handleUp = (upEvent) => finishPointerDrag(upEvent, true);
    session.handleCancel = (cancelEvent) => finishPointerDrag(cancelEvent, false);
    pointerDragSessionRef.current = session;

    try {
      event.currentTarget?.setPointerCapture?.(event.pointerId);
    } catch (error) {
      void error;
    }

    document.addEventListener('pointermove', session.handleMove, { passive: false });
    document.addEventListener('pointerup', session.handleUp, { passive: false });
    document.addEventListener('pointercancel', session.handleCancel, { passive: false });
  }, [clearPointerDragSession, dropAtPoint, positionOverlay, suppressNextClickRef]);

  useEffect(() => {
    const handleCaptureClick = (event) => {
      if (!suppressNextClickRef.current) return;
      suppressNextClickRef.current = false;
      event.preventDefault();
      event.stopPropagation();
    };

    document.addEventListener('click', handleCaptureClick, true);
    return () => document.removeEventListener('click', handleCaptureClick, true);
  }, [suppressNextClickRef]);

  useEffect(() => () => clearPointerDragSession(), [clearPointerDragSession]);

  return {
    pointerDragPreview,
    pointerDragOverlayRef,
    startPointerDrag
  };
};
