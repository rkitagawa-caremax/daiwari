import { useCallback, useLayoutEffect, useRef } from 'react';

// 最新の実装を呼びつつ、関数の識別子を恒久的に固定するフック。
// React.memo で包んだ配下 (全ページ × 全コマ) へ渡すハンドラーに使い、
// App の再レンダリングのたびに memo 境界が破れて全コマが再描画されるのを防ぐ。
// ref の更新は useLayoutEffect (描画前) で行うため、イベントが呼ぶ実装は常に最新で、
// レンダリング中の ref 書き込みも発生しない。
export const useStableHandler = (handler) => {
  const handlerRef = useRef(handler);
  useLayoutEffect(() => {
    handlerRef.current = handler;
  });
  return useCallback((...args) => handlerRef.current?.(...args), []);
};
