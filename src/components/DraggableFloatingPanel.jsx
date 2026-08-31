import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

// 画面上を自由に移動できるフローティングパネルのラッパー。
// - 子要素内の [data-drag-handle] 要素をつかんでドラッグすると移動する
// - 位置は storageKey で localStorage に保存され、次回表示時に復元される
// - 画面外にはみ出さないようにクランプする (ウィンドウリサイズ時も追従)
// - ハンドルをダブルクリックすると初期位置に戻す
const VIEWPORT_MARGIN = 4;
const INTERACTIVE_SELECTOR = 'button, input, select, textarea, a';

const readStoredPosition = (storageKey) => {
  if (!storageKey) return null;
  try {
    const raw = window.localStorage.getItem(storageKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (typeof parsed?.x !== 'number' || typeof parsed?.y !== 'number') return null;
    if (!Number.isFinite(parsed.x) || !Number.isFinite(parsed.y)) return null;
    return { x: parsed.x, y: parsed.y };
  } catch {
    return null;
  }
};

const writeStoredPosition = (storageKey, position) => {
  if (!storageKey || !position) return;
  try {
    window.localStorage.setItem(storageKey, JSON.stringify(position));
  } catch {
    // 保存できなくても動作には影響しない
  }
};

const clampPosition = (position, element) => {
  if (!element) return position;
  const rect = element.getBoundingClientRect();
  const maxX = Math.max(VIEWPORT_MARGIN, window.innerWidth - rect.width - VIEWPORT_MARGIN);
  const maxY = Math.max(VIEWPORT_MARGIN, window.innerHeight - rect.height - VIEWPORT_MARGIN);
  return {
    x: Math.min(Math.max(VIEWPORT_MARGIN, position.x), maxX),
    y: Math.min(Math.max(VIEWPORT_MARGIN, position.y), maxY)
  };
};

const findHandle = (event, container) => {
  const target = event.target instanceof Element ? event.target : null;
  if (!target) return null;
  if (target.closest(INTERACTIVE_SELECTOR)) return null; // ハンドル内のボタン等は通常操作を優先
  const handle = target.closest('[data-drag-handle]');
  if (!handle || !container?.contains(handle)) return null;
  return handle;
};

const DraggableFloatingPanel = ({
  storageKey,
  getDefaultPosition,
  className = '',
  style,
  children
}) => {
  const containerRef = useRef(null);
  const dragStateRef = useRef(null);
  const [position, setPosition] = useState(() => readStoredPosition(storageKey));
  const [isDragging, setIsDragging] = useState(false);

  const resolveDefaultPosition = useCallback(() => {
    const element = containerRef.current;
    const rect = element?.getBoundingClientRect();
    const size = { width: rect?.width || 0, height: rect?.height || 0 };
    const fallback = { x: window.innerWidth - size.width - 12, y: 80 };
    const next = typeof getDefaultPosition === 'function'
      ? getDefaultPosition({ viewportWidth: window.innerWidth, viewportHeight: window.innerHeight, ...size })
      : fallback;
    return clampPosition(next || fallback, element);
  }, [getDefaultPosition]);

  // 初回マウント時: 保存位置がなければ初期位置を計算。保存位置があっても画面内に収める。
  useLayoutEffect(() => {
    setPosition((current) => (current ? clampPosition(current, containerRef.current) : resolveDefaultPosition()));
  }, [resolveDefaultPosition]);

  // ウィンドウリサイズ時に画面内へ収め直す
  useEffect(() => {
    const handleResize = () => {
      setPosition((current) => (current ? clampPosition(current, containerRef.current) : current));
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const handlePointerDown = useCallback((event) => {
    if (event.pointerType === 'mouse' && event.button !== 0) return;
    const element = containerRef.current;
    if (!findHandle(event, element)) return;

    const rect = element.getBoundingClientRect();
    dragStateRef.current = {
      pointerId: event.pointerId,
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top
    };
    try {
      element.setPointerCapture(event.pointerId);
    } catch {
      // setPointerCapture 非対応環境でもそのまま続行
    }
    setIsDragging(true);
    event.preventDefault();
  }, []);

  const handlePointerMove = useCallback((event) => {
    const state = dragStateRef.current;
    if (!state || state.pointerId !== event.pointerId) return;
    setPosition(clampPosition(
      { x: event.clientX - state.offsetX, y: event.clientY - state.offsetY },
      containerRef.current
    ));
  }, []);

  const finishDrag = useCallback((event) => {
    const state = dragStateRef.current;
    if (!state || state.pointerId !== event.pointerId) return;
    dragStateRef.current = null;
    setIsDragging(false);
    try {
      containerRef.current?.releasePointerCapture(event.pointerId);
    } catch {
      // no-op
    }
    setPosition((current) => {
      writeStoredPosition(storageKey, current);
      return current;
    });
  }, [storageKey]);

  const handleDoubleClick = useCallback((event) => {
    if (!findHandle(event, containerRef.current)) return;
    const next = resolveDefaultPosition();
    setPosition(next);
    writeStoredPosition(storageKey, next);
  }, [resolveDefaultPosition, storageKey]);

  // ドラッグ中はテキスト選択を抑止し、カーソルを掴み状態にする
  useEffect(() => {
    if (!isDragging) return undefined;
    const previousUserSelect = document.body.style.userSelect;
    const previousCursor = document.body.style.cursor;
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'grabbing';
    return () => {
      document.body.style.userSelect = previousUserSelect;
      document.body.style.cursor = previousCursor;
    };
  }, [isDragging]);

  return (
    <div
      ref={containerRef}
      className={`fixed ${className}`}
      style={{
        left: position ? position.x : undefined,
        top: position ? position.y : undefined,
        right: position ? undefined : 12,
        visibility: position ? 'visible' : 'hidden',
        touchAction: 'none',
        ...style
      }}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={finishDrag}
      onPointerCancel={finishDrag}
      onDoubleClick={handleDoubleClick}
    >
      {children}
    </div>
  );
};

export default DraggableFloatingPanel;
