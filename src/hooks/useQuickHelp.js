import { useCallback, useEffect, useRef, useState } from 'react';

const QUICK_HELP_POPUP_HALF_WIDTH = 230;
const QUICK_HELP_BOX_SHADOW = '0 0 0 2px rgba(56, 189, 248, 0.75), 0 0 18px rgba(56, 189, 248, 0.45)';

export const useQuickHelp = () => {
  const [isQuickHelpMode, setIsQuickHelpMode] = useState(false);
  const [quickHelpPopup, setQuickHelpPopup] = useState(null);
  const quickHelpHighlightRef = useRef(null);

  const clearQuickHelpHighlight = useCallback(() => {
    const target = quickHelpHighlightRef.current;
    if (!target) return;

    target.style.boxShadow = target.dataset.quickHelpPrevBoxShadow ?? '';
    delete target.dataset.quickHelpPrevBoxShadow;
    quickHelpHighlightRef.current = null;
  }, []);

  const showQuickHelp = useCallback((event, title, description) => {
    if (!isQuickHelpMode) return;
    const target = event.currentTarget;
    if (target && quickHelpHighlightRef.current !== target) {
      clearQuickHelpHighlight();
      target.dataset.quickHelpPrevBoxShadow = target.style.boxShadow || '';
      target.style.boxShadow = QUICK_HELP_BOX_SHADOW;
      quickHelpHighlightRef.current = target;
    }

    const rect = target.getBoundingClientRect();
    const x = Math.min(
      Math.max(rect.left + rect.width / 2, QUICK_HELP_POPUP_HALF_WIDTH),
      window.innerWidth - QUICK_HELP_POPUP_HALF_WIDTH
    );
    setQuickHelpPopup({ title, description, x, y: rect.bottom + 8 });
  }, [clearQuickHelpHighlight, isQuickHelpMode]);

  const hideQuickHelp = useCallback(() => {
    clearQuickHelpHighlight();
    setQuickHelpPopup(null);
  }, [clearQuickHelpHighlight]);

  const toggleQuickHelpMode = useCallback(() => {
    if (isQuickHelpMode) hideQuickHelp();
    setIsQuickHelpMode(!isQuickHelpMode);
  }, [hideQuickHelp, isQuickHelpMode]);

  useEffect(() => clearQuickHelpHighlight, [clearQuickHelpHighlight]);

  return {
    isQuickHelpMode,
    quickHelpPopup,
    showQuickHelp,
    hideQuickHelp,
    toggleQuickHelpMode
  };
};
