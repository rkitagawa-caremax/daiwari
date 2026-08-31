import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

// 画面上を自由に移動できるフローティングパネルのラッパー。
// - 子要素内の [data-drag-handle] 要素をつかんでドラッグすると移動する
// - 位置は storageKey で localStorage に保存され、次回表示時に復元される
// - verticalResize 指定時は下端ハンドルで高さを変更し、位置と一緒に保存する
// - 画面外にはみ出さないようにクランプする (ウィンドウリサイズ時も追従)
// - ハンドルをダブルクリックすると初期位置に戻す
const VIEWPORT_MARGIN = 4;
const INTERACTIVE_SELECTOR = 'button, input, select, textarea, a';

const readStoredLayout = (storageKey) => {
  if (!storageKey) return { position: null, height: null };
  try {
    const raw = window.localStorage.getItem(storageKey);
    if (!raw) return { position: null, height: null };
    const parsed = JSON.parse(raw);
    const hasValidPosition = typeof parsed?.x === 'number'
      && typeof parsed?.y === 'number'
      && Number.isFinite(parsed.x)
      && Number.isFinite(parsed.y);
    const hasValidHeight = typeof parsed?.height === 'number'
      && Number.isFinite(parsed.height)
      && parsed.height > 0;
    return {
      position: hasValidPosition ? { x: parsed.x, y: parsed.y } : null,
      height: hasValidHeight ? parsed.height : null
    };
  } catch {
    return { position: null, height: null };
  }
};

const writeStoredLayout = (storageKey, position, height = null) => {
  if (!storageKey || !position) return;
  try {
    window.localStorage.setItem(storageKey, JSON.stringify({
      ...position,
      ...(typeof height === 'number' && Number.isFinite(height) ? { height } : {})
    }));
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

const findVerticalResizeHandle = (event, container) => {
  const target = event.target instanceof Element ? event.target : null;
  const handle = target?.closest?.('[data-vertical-resize-handle]');
  return handle && container?.contains(handle) ? handle : null;
};

const DraggableFloatingPanel = ({
  storageKey,
  getDefaultPosition,
  className = '',
  style,
  verticalResize = null,
  children
}) => {
  const isVerticalResizable = !!verticalResize;
  const resizeOptions = typeof verticalResize === 'object' && verticalResize ? verticalResize : {};
  const resizeMinHeight = Number.isFinite(resizeOptions.minHeight) ? resizeOptions.minHeight : 160;
  const resizeDefaultHeight = Number.isFinite(resizeOptions.defaultHeight) ? resizeOptions.defaultHeight : 280;
  const resizeMaxHeight = Number.isFinite(resizeOptions.maxHeight) ? resizeOptions.maxHeight : Number.POSITIVE_INFINITY;
  const [initialLayout] = useState(() => readStoredLayout(storageKey));
  const containerRef = useRef(null);
  const dragStateRef = useRef(null);
  const resizeStateRef = useRef(null);
  const [position, setPosition] = useState(initialLayout.position);
  const [height, setHeight] = useState(() => (
    isVerticalResizable ? (initialLayout.height || resizeDefaultHeight) : null
  ));
  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);

  const clampResizeHeight = useCallback((nextHeight) => {
    const elementTop = containerRef.current?.getBoundingClientRect().top ?? VIEWPORT_MARGIN;
    const viewportMaximum = Math.max(80, window.innerHeight - elementTop - VIEWPORT_MARGIN);
    const maximum = Math.max(80, Math.min(resizeMaxHeight, viewportMaximum));
    const minimum = Math.min(resizeMinHeight, maximum);
    return Math.min(Math.max(minimum, nextHeight), maximum);
  }, [resizeMaxHeight, resizeMinHeight]);

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
    const frameId = window.requestAnimationFrame(() => {
      if (isVerticalResizable) {
        setHeight((current) => clampResizeHeight(current || resizeDefaultHeight));
      }
      setPosition((current) => (current ? clampPosition(current, containerRef.current) : resolveDefaultPosition()));
    });
    return () => window.cancelAnimationFrame(frameId);
  }, [clampResizeHeight, isVerticalResizable, resizeDefaultHeight, resolveDefaultPosition]);

  // ウィンドウリサイズ時に画面内へ収め直す
  useEffect(() => {
    const handleResize = () => {
      if (isVerticalResizable) {
        setHeight((current) => clampResizeHeight(current || resizeDefaultHeight));
      }
      setPosition((current) => (current ? clampPosition(current, containerRef.current) : current));
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, [clampResizeHeight, isVerticalResizable, resizeDefaultHeight]);

  const handlePointerDown = useCallback((event) => {
    if (event.pointerType === 'mouse' && event.button !== 0) return;
    const element = containerRef.current;
    if (isVerticalResizable && findVerticalResizeHandle(event, element)) {
      const rect = element.getBoundingClientRect();
      resizeStateRef.current = {
        pointerId: event.pointerId,
        startY: event.clientY,
        startHeight: rect.height
      };
      try {
        element.setPointerCapture(event.pointerId);
      } catch {
        // setPointerCapture 非対応環境でもそのまま続行
      }
      setIsResizing(true);
      event.preventDefault();
      return;
    }
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
  }, [isVerticalResizable]);

  const handlePointerMove = useCallback((event) => {
    const resizeState = resizeStateRef.current;
    if (resizeState?.pointerId === event.pointerId) {
      setHeight(clampResizeHeight(resizeState.startHeight + event.clientY - resizeState.startY));
      return;
    }

    const state = dragStateRef.current;
    if (!state || state.pointerId !== event.pointerId) return;
    setPosition(clampPosition(
      { x: event.clientX - state.offsetX, y: event.clientY - state.offsetY },
      containerRef.current
    ));
  }, [clampResizeHeight]);

  const finishInteraction = useCallback((event) => {
    const resizeState = resizeStateRef.current;
    if (resizeState?.pointerId === event.pointerId) {
      resizeStateRef.current = null;
      setIsResizing(false);
      try {
        containerRef.current?.releasePointerCapture(event.pointerId);
      } catch {
        // no-op
      }
      setHeight((current) => {
        writeStoredLayout(storageKey, position, current);
        return current;
      });
      return;
    }

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
      writeStoredLayout(storageKey, current, height);
      return current;
    });
  }, [height, position, storageKey]);

  const handleDoubleClick = useCallback((event) => {
    if (isVerticalResizable && findVerticalResizeHandle(event, containerRef.current)) {
      const nextHeight = clampResizeHeight(resizeDefaultHeight);
      setHeight(nextHeight);
      writeStoredLayout(storageKey, position, nextHeight);
      event.preventDefault();
      return;
    }
    if (!findHandle(event, containerRef.current)) return;
    const next = resolveDefaultPosition();
    setPosition(next);
    writeStoredLayout(storageKey, next, height);
  }, [clampResizeHeight, height, isVerticalResizable, position, resizeDefaultHeight, resolveDefaultPosition, storageKey]);

  const handleResizeKeyDown = useCallback((event) => {
    if (!['ArrowUp', 'ArrowDown'].includes(event.key)) return;
    event.preventDefault();
    const delta = event.key === 'ArrowUp' ? -16 : 16;
    setHeight((current) => {
      const nextHeight = clampResizeHeight((current || resizeDefaultHeight) + delta);
      writeStoredLayout(storageKey, position, nextHeight);
      return nextHeight;
    });
  }, [clampResizeHeight, position, resizeDefaultHeight, storageKey]);

  // ドラッグ中はテキスト選択を抑止し、カーソルを掴み状態にする
  useEffect(() => {
    if (!isDragging && !isResizing) return undefined;
    const previousUserSelect = document.body.style.userSelect;
    const previousCursor = document.body.style.cursor;
    document.body.style.userSelect = 'none';
    document.body.style.cursor = isResizing ? 'ns-resize' : 'grabbing';
    return () => {
      document.body.style.userSelect = previousUserSelect;
      document.body.style.cursor = previousCursor;
    };
  }, [isDragging, isResizing]);

  return (
    <div
      ref={containerRef}
      className={`fixed ${className}`}
      style={{
        left: position ? position.x : undefined,
        top: position ? position.y : undefined,
        right: position ? undefined : 12,
        height: isVerticalResizable ? height : undefined,
        visibility: position ? 'visible' : 'hidden',
        touchAction: 'none',
        ...style
      }}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={finishInteraction}
      onPointerCancel={finishInteraction}
      onDoubleClick={handleDoubleClick}
    >
      {children}
      {isVerticalResizable && (
        <button
          type="button"
          data-vertical-resize-handle="true"
          className="absolute bottom-1 left-1/2 z-50 flex h-3 w-12 -translate-x-1/2 cursor-ns-resize items-center justify-center rounded-full border border-slate-200/80 bg-white/90 shadow-sm transition-colors hover:bg-indigo-50 focus:outline-none focus:ring-2 focus:ring-indigo-400"
          style={{ touchAction: 'none' }}
          onKeyDown={handleResizeKeyDown}
          title="上下にドラッグして高さを変更 / ダブルクリックで初期サイズ"
          aria-label="仮置き場の高さを変更"
        >
          <span className="h-0.5 w-6 rounded-full bg-slate-400" />
        </button>
      )}
    </div>
  );
};

export default DraggableFloatingPanel;
