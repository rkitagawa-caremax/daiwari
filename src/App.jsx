import { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  GoogleAuthProvider,
  browserLocalPersistence,
  setPersistence,
  signInWithPopup,
  signInWithRedirect,
  getRedirectResult,
  onAuthStateChanged,
  signOut
} from 'firebase/auth';
import {
  collection,
  query,
  where,
  doc,
  setDoc,
  addDoc,
  onSnapshot,
  deleteDoc,
  updateDoc,
  serverTimestamp,
  writeBatch,
  getDocs,
  runTransaction,
  deleteField,
  increment,
  arrayUnion
} from 'firebase/firestore';
import {
  ChevronLeft,
  ChevronRight,
  Layout,
  Check,
  Maximize,
  Copy,
  Loader2,
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  MoreHorizontal
} from 'lucide-react';

import { idbHelper } from './idbHelper';
import {
  GOOGLE_ALLOWED_ACCOUNTS,
  buildGoogleAuthErrorMessage,
  expandAllowedEmailVariants,
  isAllowedGoogleUser,
  normalizeEmail
} from './config/authPolicy';
import {
  USE_LOCAL_STORAGE,
  DEFAULT_APP_ID,
  auth,
  db,
  CLOUD_IMAGES_CACHE_KEY,
  CLOUD_SALES_CACHE_KEY,
  CLOUD_CACHE_TTL_MS,
  LOCAL_WORK_LOGS_KEY
} from './config/firebase';
import {
  GENRES
} from './constants/layout';
import {
  DEFAULT_PANEL_DATA,
  PANEL_COUNT,
  applyPanelTransferableContent,
  buildDefaultPanels,
  buildPanelMapUpdates,
  clearPanelTransferableContent,
  getPanelDataPatch,
  getPanelFreeLabels,
  getPanelsFromDocData,
  getPanelTransferableContent,
  hasPanelTransferableContent,
  sanitizePanelData,
  toPanelsMap
} from './domain/panels';
import {
  buildPanelArrangeFinalPanels,
  buildPanelArrangeView,
  createPanelArrangeSession,
  getUnresolvedPanelArrangeTokens,
  hasPanelImageContent,
  reconcilePanelArrangeSession,
  stagePanelArrangeDrop
} from './domain/panelArrange';
import {
  isSameStockImageList,
  normalizeStockImages
} from './domain/images';
import { buildPdfExportPlan } from './domain/pdfExport';
import { WORK_ACTIONS, applyWorkLogDeltaToRecord } from './domain/workActivity';
import {
  isSameSheetList,
  isSameTransferItemList,
  toComparableSeconds
} from './domain/workspaceComparators';
import { normalizeCode } from './domain/productCodes';
import {
  getCoords,
  getSizeType
} from './domain/panelLayout';
import {
  DAIWARI_DROPZONE_ATTR,
  DAIWARI_PANEL_DROPZONE_PREFIX,
  PANEL_ARRANGE_HOLD_MS,
  POINTER_DRAG_THRESHOLD_PX,
  clearActiveNativeDragPayload,
  extractPanelAssignmentFromDragPayload,
  extractPanelArrangeDragPayload,
  extractPanelMoveDragPayload,
  getActiveNativeDragPayload,
  getActivePanelMoveDragPayload,
  getDragPayload,
  hasPanelArrangeHoldMoved,
  isDropEventHandled,
  markDropEventHandled,
  normalizeDragPayload,
  parseNullableDragValue
} from './lib/dragPayload';
import { parseCSVLine, readFileAutoEncoding } from './lib/csv';
import { downloadTextFile } from './lib/download';
import {
  buildDatedCsvFilename,
  buildExcludedItemsCsvContent,
  buildPageCsvContent
} from './domain/pageCsv';
import {
  buildImportedSheets,
  buildPageCsvImportReport,
  parsePageCsvRows
} from './domain/pageCsvImport';
import { createPdfRenderer, waitForPdfExportSurface } from './lib/pdfExport';
import { useWorkActivityTracker } from './hooks/useWorkActivityTracker';
import { useWorkspaceUndoState } from './hooks/useWorkspaceUndoState';
import {
  buildFirestoreActionErrorMessage,
  getFirestoreErrorCode,
  retryAsync
} from './lib/firestoreErrors';
import { compressImage } from './lib/imageProcessing';
import AlertModal from './components/dialogs/AlertModal';
import ConfirmModal from './components/dialogs/ConfirmModal';
import HiddenImportModal from './components/dialogs/HiddenImportModal';
import ImagePreviewModal from './components/dialogs/ImagePreviewModal';
import ProcessingModal from './components/dialogs/ProcessingModal';
import SettingsModal from './components/dialogs/SettingsModal';
import AuthGate from './features/auth/AuthGate';
import SalesCodeLookupModal from './features/sales/SalesCodeLookupModal';
import SalesPopup from './features/sales/SalesPopup';
import Sheet from './features/sheets/components/Sheet';
import SheetControlPanel from './features/sheets/components/SheetControlPanel';
import PdfExportSurface from './features/sheets/components/PdfExportSurface';
import Sidebar from './features/sidebar/Sidebar';
import TempShelfPanel from './features/sidebar/TempShelfPanel';
import DraggableFloatingPanel from './components/DraggableFloatingPanel';
import QuickHelpPopup from './features/layout/QuickHelpPopup';
import UndoNoticeToast from './features/layout/UndoNoticeToast';
import PointerDragPreview from './features/layout/PointerDragPreview';
import PanelArrangeBanner from './features/layout/PanelArrangeBanner';
import TopBarsToggleButton from './features/layout/TopBarsToggleButton';
import ZoomControls from './features/layout/ZoomControls';
import PageSelectionToolbar from './features/layout/PageSelectionToolbar';
import HeaderToolsMenu from './features/layout/HeaderToolsMenu';
import ContentHeaderControls from './features/layout/ContentHeaderControls';
import AppHeader from './features/layout/AppHeader';

// フローティングパネルの初期位置 (右端寄せ)。従来の「右端・縦中央付近に縦積み」を再現する。
const FLOATING_PANEL_RIGHT_MARGIN = 12;
const FLOATING_PANEL_GAP = 8;
const SHEET_CONTROL_PANEL_ESTIMATED_HEIGHT = 168;
const getFloatingPanelStackTop = (viewportHeight) => Math.max(72, Math.round(viewportHeight * 0.5) - 240);
const getSheetControlPanelDefaultPosition = ({ viewportWidth, viewportHeight, width }) => ({
  x: viewportWidth - width - FLOATING_PANEL_RIGHT_MARGIN,
  y: getFloatingPanelStackTop(viewportHeight)
});
const getTempShelfPanelDefaultPosition = ({ viewportWidth, viewportHeight, width }) => ({
  x: viewportWidth - width - FLOATING_PANEL_RIGHT_MARGIN,
  y: getFloatingPanelStackTop(viewportHeight) + SHEET_CONTROL_PANEL_ESTIMATED_HEIGHT + FLOATING_PANEL_GAP
});
import WorkLogDashboard from './features/workLogs/WorkLogDashboard';
import {
  applyUndoEntryToWorkspace,
  countUndoEntryChanges,
  invertUndoEntry,
  restoreCloudUndoEntry,
  undoEntryHasClientConflict
} from './features/undo/accountUndo';

// --- Components ---

export default function App() {
  // 初期表示のAppIDを決定（URLパラメータ > LocalStorage > Default）
  const getInitialAppId = () => {
    const params = new URLSearchParams(window.location.search);
    const urlProject = params.get('project');
    if (urlProject) return urlProject;

    const saved = localStorage.getItem('daiwari_active_workspace');
    return saved || DEFAULT_APP_ID;
  };

  const initialAppId = getInitialAppId();
  const [appId] = useState(initialAppId);
  const [firebaseUser, setFirebaseUser] = useState(null);
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isAuthReady, setIsAuthReady] = useState(USE_LOCAL_STORAGE);
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [authErrorMessage, setAuthErrorMessage] = useState('');

  useEffect(() => {
    try {
      localStorage.setItem('daiwari_active_workspace', appId);
      const url = new URL(window.location);
      if (url.searchParams.get('project') !== appId) {
        url.searchParams.set('project', appId);
        window.history.replaceState({}, '', url);
      }
    } catch (e) {
      void e;
    }
  }, [appId]);

  // Data State
  const undoAccountId = USE_LOCAL_STORAGE ? 'local_user' : (firebaseUser?.uid || null);
  const {
    excludedItems,
    getLatestRedoEntry,
    getLatestUndoEntry,
    images,
    isUndoApplyingRef,
    isUndoBusyRef,
    pushRedoEntry,
    pushUndoEntry,
    removeRedoEntry,
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
  } = useWorkspaceUndoState({ accountId: undoAccountId });
  const [salesData, setSalesData] = useState(null); // { code: [{name, spec, count}] }
  const [salesDataLastUpdated, setSalesDataLastUpdated] = useState(null);

  // UI State
  const [viewMode, setViewMode] = useState('overview');
  const [zoomScale, setZoomScale] = useState(1);
  const [activeSheetId, setActiveSheetId] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [isTopBarsVisible, setIsTopBarsVisible] = useState(true);
  const [isToolsMenuOpen, setIsToolsMenuOpen] = useState(false);
  const [sidebarWidth, setSidebarWidth] = useState(280);
  const [searchQuery, setSearchQuery] = useState("");
  const [genreFilter, setGenreFilter] = useState('all');
  const [selection, setSelection] = useState({ sheetId: null, indices: [] });
  const [isMergeMode, setIsMergeMode] = useState(false);
  const [highlightEmpty, setHighlightEmpty] = useState(false);
  const [highlightLabels, setHighlightLabels] = useState(false);
  const [confirmDialog, setConfirmDialog] = useState({ isOpen: false, message: '', onConfirm: null });
  const [alertDialog, setAlertDialog] = useState({ isOpen: false, message: '', title: '通知', closeOnBackdrop: false });
  const fileInputRef = useRef(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isHiddenImportModalOpen, setIsHiddenImportModalOpen] = useState(false);
  const [isWorkLogDashboardOpen, setIsWorkLogDashboardOpen] = useState(false);
  const [workLogRecords, setWorkLogRecords] = useState([]);
  const [isWorkLogLoading, setIsWorkLogLoading] = useState(false);
  const [workLogErrorMessage, setWorkLogErrorMessage] = useState('');
  const [isQuickHelpMode, setIsQuickHelpMode] = useState(false);
  const [quickHelpPopup, setQuickHelpPopup] = useState(null);
  const logoTapCountRef = useRef(0);
  const logoTapTimeoutRef = useRef(null);
  const quickHelpHighlightRef = useRef(null);
  const [isSalesMode, setIsSalesMode] = useState(false); // 実績モード
  const [isSalesLookupOpen, setIsSalesLookupOpen] = useState(false);
  const [isLabelSelectionMode, setIsLabelSelectionMode] = useState(false);
  const [pointerDragPreview, setPointerDragPreview] = useState(null);
  const [panelArrangeSession, setPanelArrangeSession] = useState(null);
  const [arrangeDraggingTokenId, setArrangeDraggingTokenId] = useState(null);
  const [isPanelArrangeFinalizing, setIsPanelArrangeFinalizing] = useState(false);
  const panelArrangeModeSheetId = panelArrangeSession?.sheetId || null;
  const [assignedImagePreview, setAssignedImagePreview] = useState(null);
  const [undoNotice, setUndoNotice] = useState(null);
  const undoNoticeTimerRef = useRef(null);
  const salesModeLongPressTimerRef = useRef(null);
  const salesModeLongPressTriggeredRef = useRef(false);
  const pointerDragOverlayRef = useRef(null);
  const pointerDragSessionRef = useRef(null);
  const panelArrangeHoldRef = useRef(null);
  const suppressNextClickRef = useRef(false);

  // Sales Popup State
  const [hoveredSalesData, setHoveredSalesData] = useState(null);
  const [salesPopupPos, setSalesPopupPos] = useState(null);
  const closeTimeoutRef = useRef(null);

  const [isPageSelectionMode, setIsPageSelectionMode] = useState(false);
  const [selectedSheetIds, setSelectedSheetIds] = useState(new Set());

  // === 画面ロック (鍵ボタン 2秒長押しでトグル) ===
  // ロック中: 編集系 (panel 更新 / DnD 配置 / シート追加削除 / 画像管理 / CSV 取り込み / 結合・分離 / 仮置き場 / 除外 等) を一律 no-op
  // ロック中も可能: viewMode 切替 / 実績モード / ページ移動 / 検索 / Sidebar 閲覧 / プレビュー
  const [isLocked, setIsLocked] = useState(false);
  const isLockedRef = useRef(false);
  useEffect(() => { isLockedRef.current = isLocked; }, [isLocked]);

  useEffect(() => () => {
    if (undoNoticeTimerRef.current) clearTimeout(undoNoticeTimerRef.current);
  }, []);

  const lockHoldTimerRef = useRef(null);
  const lockHoldFiredRef = useRef(false);
  const LOCK_HOLD_MS = 2000;

  const startLockHold = useCallback(() => {
    if (panelArrangeModeSheetId) return;
    if (lockHoldTimerRef.current) clearTimeout(lockHoldTimerRef.current);
    lockHoldFiredRef.current = false;
    lockHoldTimerRef.current = setTimeout(() => {
      lockHoldFiredRef.current = true;
      setIsLocked((prev) => !prev);
      lockHoldTimerRef.current = null;
    }, LOCK_HOLD_MS);
  }, [panelArrangeModeSheetId]);

  const cancelLockHold = useCallback(() => {
    if (lockHoldTimerRef.current) {
      clearTimeout(lockHoldTimerRef.current);
      lockHoldTimerRef.current = null;
    }
  }, []);

  useEffect(() => () => {
    if (lockHoldTimerRef.current) clearTimeout(lockHoldTimerRef.current);
  }, []);

  const clearPanelArrangeHold = useCallback(() => {
    const session = panelArrangeHoldRef.current;
    if (!session) return;

    if (session.timerId) clearTimeout(session.timerId);
    if (typeof document !== 'undefined') {
      document.removeEventListener('pointermove', session.handleMove);
      document.removeEventListener('pointerup', session.handleRelease);
      document.removeEventListener('pointercancel', session.handleRelease);
      document.removeEventListener('scroll', session.handleRelease, true);
    }
    if (typeof window !== 'undefined') {
      window.removeEventListener('blur', session.handleRelease);
    }
    panelArrangeHoldRef.current = null;
  }, []);

  const clearPanelArrangeModeState = useCallback(() => {
    clearPanelArrangeHold();
    setArrangeDraggingTokenId(null);
    setPanelArrangeSession(null);
    setIsPanelArrangeFinalizing(false);
  }, [clearPanelArrangeHold]);

  const startPanelArrangeHold = useCallback((event, target = {}) => {
    if (
      !event
      || event.isPrimary === false
      || (event.button !== undefined && event.button !== 0)
      || viewMode !== 'single'
      || !activeSheetId
      || target.sheetId !== activeSheetId
      || panelArrangeModeSheetId
      || isLocked
      || isPageSelectionMode
      || isLabelSelectionMode
      || isSalesMode
    ) return;

    clearPanelArrangeHold();

    const session = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      timerId: null,
      handleMove: null,
      handleRelease: null
    };

    session.handleMove = (moveEvent) => {
      if (!moveEvent || moveEvent.pointerId !== session.pointerId) return;
      if (hasPanelArrangeHoldMoved(
        session.startX,
        session.startY,
        moveEvent.clientX,
        moveEvent.clientY
      )) {
        clearPanelArrangeHold();
      }
    };
    session.handleRelease = (releaseEvent) => {
      if (releaseEvent?.pointerId !== undefined && releaseEvent.pointerId !== session.pointerId) return;
      clearPanelArrangeHold();
    };
    session.timerId = setTimeout(() => {
      if (panelArrangeHoldRef.current !== session) return;
      clearPanelArrangeHold();
      const targetSheet = sheets.find((sheet) => sheet.id === target.sheetId);
      if (!targetSheet) return;
      const targetPanel = targetSheet.panels?.[target.panelIndex];
      if (!hasPanelImageContent(targetPanel)) return;
      const arrangeSession = createPanelArrangeSession(target.sheetId, targetSheet.panels || []);
      if (arrangeSession.tokens.length === 0) return;
      suppressNextClickRef.current = true;
      setArrangeDraggingTokenId(null);
      setIsLabelSelectionMode(false);
      setIsSalesMode(false);
      setPanelArrangeSession(arrangeSession);
      try {
        navigator.vibrate?.(35);
      } catch (error) {
        void error;
      }
    }, PANEL_ARRANGE_HOLD_MS);

    panelArrangeHoldRef.current = session;
    document.addEventListener('pointermove', session.handleMove, { passive: true });
    document.addEventListener('pointerup', session.handleRelease);
    document.addEventListener('pointercancel', session.handleRelease);
    document.addEventListener('scroll', session.handleRelease, true);
    window.addEventListener('blur', session.handleRelease);
  }, [
    activeSheetId,
    clearPanelArrangeHold,
    isLabelSelectionMode,
    isLocked,
    isPageSelectionMode,
    isSalesMode,
    panelArrangeModeSheetId,
    sheets,
    viewMode
  ]);

  const handleArrangeDragStateChange = useCallback((tokenId, isDragging) => {
    setArrangeDraggingTokenId((current) => {
      if (!isDragging) return current === tokenId ? null : current;
      return panelArrangeSession?.tokens?.some((token) => token.id === tokenId) ? tokenId : current;
    });
  }, [panelArrangeSession]);

  useEffect(() => () => {
    clearPanelArrangeHold();
  }, [clearPanelArrangeHold]);

  useEffect(() => {
    if (!panelArrangeModeSheetId) return;
    if (viewMode !== 'single') setViewMode('single');
    if (activeSheetId !== panelArrangeModeSheetId) setActiveSheetId(panelArrangeModeSheetId);
    if (isPageSelectionMode) setIsPageSelectionMode(false);
    if (isLabelSelectionMode) setIsLabelSelectionMode(false);
    if (isSalesMode) setIsSalesMode(false);
  }, [
    activeSheetId,
    isLabelSelectionMode,
    isPageSelectionMode,
    isSalesMode,
    panelArrangeModeSheetId,
    viewMode
  ]);

  const [isProcessing, setIsProcessing] = useState(false);
  const [progressValue, setProgressValue] = useState(0);
  const [progressMax, setProgressMax] = useState(100);
  const [progressMessage, setProgressMessage] = useState("");
  const [isDataLoaded, setIsDataLoaded] = useState(false); // データ読み込み完了フラグ
  const [useLegacyTempShelf, setUseLegacyTempShelf] = useState(false);
  const [pdfExportPage, setPdfExportPage] = useState(null);

  // コレクション参照を appId に依存させる
  const sheetsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'sheets'), [appId, db]);
  const imagesCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'images'), [appId, db]);
  const tempShelfUserId = useMemo(() => USE_LOCAL_STORAGE ? 'local' : (firebaseUser?.uid || null), [firebaseUser]);
  const userTempShelfCollection = useMemo(() => {
    if (USE_LOCAL_STORAGE) return null;
    if (!tempShelfUserId) return null;
    return collection(db, 'artifacts', appId, 'users', tempShelfUserId, 'tempShelf');
  }, [appId, db, tempShelfUserId]);
  const legacyTempShelfCollection = useMemo(() => {
    if (USE_LOCAL_STORAGE) return null;
    return collection(db, 'artifacts', appId, 'public', 'data', 'tempShelf');
  }, [appId, db]);
  const tempShelfCollection = useMemo(() => {
    if (USE_LOCAL_STORAGE) return null;
    return useLegacyTempShelf ? legacyTempShelfCollection : userTempShelfCollection;
  }, [useLegacyTempShelf, legacyTempShelfCollection, userTempShelfCollection]);
  const tempShelfSyncSource = useMemo(() => {
    if (USE_LOCAL_STORAGE || !tempShelfCollection) return null;
    if (useLegacyTempShelf && tempShelfUserId) {
      return query(tempShelfCollection, where('ownerUid', '==', tempShelfUserId));
    }
    return tempShelfCollection;
  }, [tempShelfCollection, useLegacyTempShelf, tempShelfUserId]);
  const excludedItemsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'excludedItems'), [appId, db]);
  const settingsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'settings'), [appId, db]);
  const salesChunksCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'salesDataChunks'), [appId, db]);
  const workLogsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'activityLogs'), [appId]);
  const signedInUserName = useMemo(() => {
    const signedInEmail = normalizeEmail(firebaseUser?.email);
    if (signedInEmail) {
      const matchedAccount = GOOGLE_ALLOWED_ACCOUNTS.find((account) =>
        expandAllowedEmailVariants(account.email).includes(signedInEmail)
      );
      if (matchedAccount?.name) return matchedAccount.name;
    }
    return firebaseUser?.displayName || firebaseUser?.email || '';
  }, [firebaseUser]);

  // --- Auth & Init ---
  useEffect(() => {
    // ローカルストレージモードの場合はFirebase認証をスキップ
    if (USE_LOCAL_STORAGE) {
      setFirebaseUser({ uid: 'local_user' }); // ダミーユーザー
      setIsAuthenticated(true);
      setIsAuthReady(true);

      const bootApp = async () => {
        try {
          // IndexedDBを優先
          let savedSheets = await idbHelper.getItem('sheets');
          let savedImages = await idbHelper.getItem('images');
          let savedTempItems = await idbHelper.getItem('tempItems');
          let savedExcludedItems = await idbHelper.getItem('excludedItems');
          let savedSalesData = await idbHelper.getItem('salesData');

          // 初回アクセス時のみLocalStorageからの移行を試みる
          const lsFlag = localStorage.getItem('daiwari_migrated_to_idb');
          if (!lsFlag) {
            console.log("Checking for localStorage data to migrate...");
            const lsSheets = localStorage.getItem('daiwari_sheets');
            if (lsSheets && !savedSheets) {
              console.log("Migrating sheets...");
              savedSheets = JSON.parse(lsSheets);
              await idbHelper.setItem('sheets', savedSheets);
            }
            const lsImages = localStorage.getItem('daiwari_images');
            if (lsImages && !savedImages) {
              console.log("Migrating images...");
              savedImages = JSON.parse(lsImages);
              await idbHelper.setItem('images', savedImages);
            }
            if (!savedTempItems) {
              savedTempItems = JSON.parse(localStorage.getItem('daiwari_tempItems') || '[]');
              await idbHelper.setItem('tempItems', savedTempItems);
            }
            if (!savedExcludedItems) {
              savedExcludedItems = JSON.parse(localStorage.getItem('daiwari_excludedItems') || '[]');
              await idbHelper.setItem('excludedItems', savedExcludedItems);
            }
            if (!savedSalesData) {
              savedSalesData = JSON.parse(localStorage.getItem('daiwari_salesData') || 'null');
              if (savedSalesData) await idbHelper.setItem('salesData', savedSalesData);
            }
            localStorage.setItem('daiwari_migrated_to_idb', 'true');
          }

          const loadedImages = Array.isArray(savedImages) ? savedImages : [];
          const loadedImageDataById = {};
          loadedImages.forEach((img) => {
            if (img?.id && (img?.data || img?.image)) {
              loadedImageDataById[img.id] = img.data || img.image;
            }
          });
          const normalizedSavedImages = normalizeStockImages(loadedImages, loadedImageDataById);

          syncSheets(savedSheets || []);
          syncImages(normalizedSavedImages);
          syncTempItems(savedTempItems || []);
          syncExcludedItems(savedExcludedItems || []);
          if (savedSalesData) setSalesData(savedSalesData);

          if (!isSameStockImageList(loadedImages, normalizedSavedImages)) {
            await idbHelper.setItem('images', normalizedSavedImages);
          }

          setIsDataLoaded(true);
        } catch (err) {
          console.error("Initialization failed:", err);
          setIsDataLoaded(true);
        }
      };
      bootApp();
      return;
    }

    const initialize = async () => {
      const host = typeof window !== 'undefined' ? window.location.hostname : '';
      if (host === '127.0.0.1') {
        try {
          const redirected = new URL(window.location.href);
          redirected.hostname = 'localhost';
          window.location.replace(redirected.toString());
          return;
        } catch (error) {
          console.error('Failed to normalize host to localhost:', error);
        }
      }

      try {
        await setPersistence(auth, browserLocalPersistence);
        try {
          await getRedirectResult(auth);
        } catch (redirectError) {
          console.error('Google redirect result failed:', redirectError);
          setAuthErrorMessage(buildGoogleAuthErrorMessage(redirectError, 'Googleログイン結果の復元に失敗しました。再度お試しください。'));
        }
      } catch (error) {
        console.error("Auth initialization failed:", error);
        setAuthErrorMessage(buildGoogleAuthErrorMessage(error, '認証初期化に失敗しました。再度お試しください。'));
      } finally {
        setIsAuthReady(true);
      }
    };
    initialize();

    const unsubscribe = onAuthStateChanged(auth, async (u) => {
      if (!u) {
        setIsAuthReady(true);
        setFirebaseUser(null);
        setIsAuthenticated(false);
        return;
      }

      const allowed = isAllowedGoogleUser(u);
      if (allowed) {
        setIsAuthReady(true);
        setFirebaseUser(u);
        setIsAuthenticated(true);
        setAuthErrorMessage('');
        return;
      }

      setIsAuthReady(true);
      setFirebaseUser(null);
      setIsAuthenticated(false);
      const signedInEmail = normalizeEmail(u?.email);
      setAuthErrorMessage(
        signedInEmail
          ? `許可対象外のアカウントです: ${signedInEmail}`
          : '許可されているGoogleアカウントでログインしてください。'
      );
      try {
        await signOut(auth);
      } catch (error) {
        console.error('Sign out unauthorized account failed:', error);
      }
    });
    return () => unsubscribe();
  }, [syncExcludedItems, syncImages, syncSheets, syncTempItems]);

  const handleGoogleSignIn = useCallback(async () => {
    if (USE_LOCAL_STORAGE || !auth) return;
    setIsSigningIn(true);
    setAuthErrorMessage('');
    try {
      const provider = new GoogleAuthProvider();
      provider.setCustomParameters({ prompt: 'select_account' });
      await signInWithPopup(auth, provider);
    } catch (error) {
      const code = String(error?.code || '').toLowerCase();
      if (code === 'auth/popup-blocked' || code === 'auth/operation-not-supported-in-this-environment') {
        try {
          const provider = new GoogleAuthProvider();
          provider.setCustomParameters({ prompt: 'select_account' });
          await signInWithRedirect(auth, provider);
          return;
        } catch (redirectError) {
          setAuthErrorMessage(buildGoogleAuthErrorMessage(redirectError));
          console.error('Google redirect sign-in failed:', redirectError);
        }
      } else {
        setAuthErrorMessage(buildGoogleAuthErrorMessage(error));
      }
      console.error('Google sign-in failed:', error);
    } finally {
      setIsSigningIn(false);
    }
  }, []);

  const handleLogout = useCallback(async () => {
    if (USE_LOCAL_STORAGE || !auth) return;
    try {
      await signOut(auth);
      setIsAuthenticated(false);
      setFirebaseUser(null);
      syncSheets([]);
      syncImages([]);
      syncTempItems([]);
      syncExcludedItems([]);
      setSalesData(null);
      setIsDataLoaded(false);
      setAuthErrorMessage('');
    } catch (error) {
      console.error('Logout failed:', error);
    }
  }, [syncExcludedItems, syncImages, syncSheets, syncTempItems]);

  // --- Data Sync ---
  // 自動保存 (Auto-Save) - IndexedDB with Debounce
  const saveTimeoutRef = useRef(null);
  const localSavedSnapshotRef = useRef({
    sheets: null,
    images: null,
    tempItems: null,
    excludedItems: null,
    salesData: null
  });

  useEffect(() => {
    if (!(USE_LOCAL_STORAGE && isDataLoaded)) {
      return () => {
        if (saveTimeoutRef.current) {
          clearTimeout(saveTimeoutRef.current);
        }
      };
    }

    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    const shouldSaveSheets = localSavedSnapshotRef.current.sheets !== sheets;
    const shouldSaveImages = localSavedSnapshotRef.current.images !== images;
    const shouldSaveTempItems = localSavedSnapshotRef.current.tempItems !== tempItems;
    const shouldSaveExcludedItems = localSavedSnapshotRef.current.excludedItems !== excludedItems;
    const shouldSaveSalesData = localSavedSnapshotRef.current.salesData !== salesData;

    if (!shouldSaveSheets && !shouldSaveImages && !shouldSaveTempItems && !shouldSaveExcludedItems && !shouldSaveSalesData) {
      return () => {
        if (saveTimeoutRef.current) {
          clearTimeout(saveTimeoutRef.current);
        }
      };
    }

    const snapshot = { sheets, images, tempItems, excludedItems, salesData };
    saveTimeoutRef.current = setTimeout(async () => {
      try {
        const tasks = [];
        if (shouldSaveSheets) {
          tasks.push(idbHelper.setItem('sheets', snapshot.sheets).then(() => {
            localSavedSnapshotRef.current.sheets = snapshot.sheets;
          }));
        }
        if (shouldSaveImages) {
          tasks.push(idbHelper.setItem('images', snapshot.images).then(() => {
            localSavedSnapshotRef.current.images = snapshot.images;
          }));
        }
        if (shouldSaveTempItems) {
          tasks.push(idbHelper.setItem('tempItems', snapshot.tempItems).then(() => {
            localSavedSnapshotRef.current.tempItems = snapshot.tempItems;
          }));
        }
        if (shouldSaveExcludedItems) {
          tasks.push(idbHelper.setItem('excludedItems', snapshot.excludedItems).then(() => {
            localSavedSnapshotRef.current.excludedItems = snapshot.excludedItems;
          }));
        }
        if (shouldSaveSalesData && snapshot.salesData) {
          tasks.push(idbHelper.setItem('salesData', snapshot.salesData).then(() => {
            localSavedSnapshotRef.current.salesData = snapshot.salesData;
          }));
        }
        await Promise.all(tasks);
      } catch (err) {
        console.error("Auto-save failed:", err);
      }
    }, 500);

    return () => {
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current);
      }
    };
  }, [sheets, images, tempItems, excludedItems, salesData, isDataLoaded]);

  // 全体表示に切り替えた時、実績モードを自動的にオフにする
  useEffect(() => {
    if (viewMode !== 'list' && viewMode !== 'single' && isSalesMode) {
      setIsSalesMode(false);
    }
  }, [viewMode]);

  useEffect(() => {
    return () => {
      if (logoTapTimeoutRef.current) {
        clearTimeout(logoTapTimeoutRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!isDataLoaded || images.length === 0) return;

    const imageDataMap = {};
    images.forEach((img) => {
      if (img?.id && (img?.data || img?.image)) {
        imageDataMap[img.id] = img.data || img.image;
      }
    });

    const normalizedImages = normalizeStockImages(images, imageDataMap);
    if (!isSameStockImageList(images, normalizedImages)) {
      syncImages(normalizedImages);
    }
  }, [images, isDataLoaded, syncImages]);

  useEffect(() => {
    // 詳細単一表示以外ではラベル配置モードを自動解除
    if (viewMode !== 'single' && isLabelSelectionMode) {
      setIsLabelSelectionMode(false);
    }
  }, [viewMode, isLabelSelectionMode]);

  useEffect(() => {
    if (viewMode === 'overview') return;
    setIsPageSelectionMode((prev) => (prev ? false : prev));
    setSelectedSheetIds((prev) => (prev.size > 0 ? new Set() : prev));
  }, [viewMode]);

  useEffect(() => {
    // 単一表示時に対象ページIDが不整合なら先頭ページへ補正
    if (viewMode !== 'single' || sheets.length === 0) return;
    const exists = sheets.some((sheet) => sheet.id === activeSheetId);
    if (!exists) {
      setActiveSheetId(sheets[0].id);
      setIsLabelSelectionMode(false);
    }
  }, [viewMode, sheets, activeSheetId]);


  useEffect(() => {
    // ローカルストレージモードの場合はFirebase同期をスキップ
    if (USE_LOCAL_STORAGE) {
      return; // データは Auth init で既に読み込み済み
    }

    if (!isAuthenticated) return;
    if (!sheetsCollection || !imagesCollection || !excludedItemsCollection || !settingsCollection) return;
    let isCancelled = false;

    const loadImagesWithCache = async () => {
      let cachedBundle = null;
      try {
        cachedBundle = await idbHelper.getItem(CLOUD_IMAGES_CACHE_KEY);
        if (isCancelled) return;
        if (Array.isArray(cachedBundle?.items)) {
          const cachedImages = normalizeStockImages(cachedBundle.items);
          syncImages((prev) => (isSameStockImageList(prev, cachedImages) ? prev : cachedImages));
        }
      } catch (error) {
        console.error("Cloud image cache load failed:", error);
      }

      const fetchedAt = Number(cachedBundle?.fetchedAt || 0);
      if (fetchedAt > 0 && Date.now() - fetchedAt < CLOUD_CACHE_TTL_MS) return;

      try {
        const snapshot = await getDocs(imagesCollection);
        if (isCancelled) return;
        const loadedImages = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
        const loadedImageDataById = {};
        loadedImages.forEach((img) => {
          if (img?.id && (img?.data || img?.image)) {
            loadedImageDataById[img.id] = img.data || img.image;
          }
        });
        const normalizedLoadedImages = normalizeStockImages(loadedImages, loadedImageDataById);
        syncImages((prev) => (isSameStockImageList(prev, normalizedLoadedImages) ? prev : normalizedLoadedImages));
        await idbHelper.setItem(CLOUD_IMAGES_CACHE_KEY, {
          items: normalizedLoadedImages,
          fetchedAt: Date.now()
        });
      } catch (err) {
        console.error("Image Load Error", err);
      }
    };

    const loadSalesWithCache = async (metaData = null) => {
      const metaSeconds = toComparableSeconds(metaData?.updatedAt);
      let cachedBundle = null;
      try {
        cachedBundle = await idbHelper.getItem(CLOUD_SALES_CACHE_KEY);
        if (isCancelled) return;
        if (cachedBundle?.data) {
          setSalesData(cachedBundle.data);
          if (!metaSeconds || cachedBundle.metaSeconds === metaSeconds) return;
        }
      } catch (error) {
        console.error("Cloud sales cache load failed:", error);
      }

      if (!salesChunksCollection) return;
      try {
        const snapshot = await getDocs(salesChunksCollection);
        if (isCancelled) return;
        const fullSalesMap = {};
        snapshot.docs.forEach(doc => {
          const data = doc.data();
          if (data.items) {
            try {
              const chunkMap = JSON.parse(data.items);
              Object.assign(fullSalesMap, chunkMap);
            } catch (e) {
              console.error("Failed to parse sales chunk", e);
            }
          }
        });
        setSalesData(fullSalesMap);
        await idbHelper.setItem(CLOUD_SALES_CACHE_KEY, {
          data: fullSalesMap,
          metaSeconds,
          fetchedAt: Date.now()
        });
      } catch (err) {
        console.error("Sales Data Load Error", err);
      }
    };

    const unsubscribeSheets = onSnapshot(sheetsCollection, (snapshot) => {
      const loadedSheets = snapshot.docs.map((snap) => {
        const data = snap.data() || {};
        return {
          ...data,
          id: snap.id,
          _hasLegacyPanels: Array.isArray(data.panels),
          panels: getPanelsFromDocData(data)
        };
      });
      loadedSheets.sort((a, b) => (a.createdAt?.seconds || 0) - (b.createdAt?.seconds || 0));
      const updateSheets = snapshot.metadata.hasPendingWrites && !isUndoApplyingRef.current
        ? setSheets
        : syncSheets;
      updateSheets((prev) => (isSameSheetList(prev, loadedSheets) ? prev : loadedSheets));
    }, (err) => console.error("Sheet Sync Error", err));

    void loadImagesWithCache();

    const unsubscribeExcluded = onSnapshot(excludedItemsCollection, (snapshot) => {
      const loadedExcluded = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
      loadedExcluded.sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      const updateExcludedItems = snapshot.metadata.hasPendingWrites && !isUndoApplyingRef.current
        ? setExcludedItems
        : syncExcludedItems;
      updateExcludedItems((prev) => (isSameTransferItemList(prev, loadedExcluded) ? prev : loadedExcluded));
    }, (err) => console.error("Excluded Items Sync Error", err));

    void loadSalesWithCache();

    const unsubscribeMeta = onSnapshot(doc(settingsCollection, 'salesDataMeta'), (docSnap) => {
      if (docSnap.exists()) {
        const metaData = docSnap.data() || {};
        setSalesDataLastUpdated(metaData.updatedAt?.toDate() || null);
        void loadSalesWithCache(metaData);
      }
    }, (err) => console.error("Sales Meta Sync Error", err));

    return () => {
      isCancelled = true;
      unsubscribeSheets();
      unsubscribeExcluded();
      unsubscribeMeta();
    };
  }, [
    excludedItemsCollection,
    imagesCollection,
    isAuthenticated,
    salesChunksCollection,
    setExcludedItems,
    setSheets,
    settingsCollection,
    sheetsCollection,
    syncExcludedItems,
    syncImages,
    syncSheets
  ]);

  useEffect(() => {
    if (USE_LOCAL_STORAGE) return;
    if (!isAuthenticated || !tempShelfCollection || !tempShelfSyncSource) return;

    const unsubscribeTemp = onSnapshot(tempShelfSyncSource, (snapshot) => {
      let loadedTemps = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
      if (useLegacyTempShelf && tempShelfUserId) {
        loadedTemps = loadedTemps.filter((item) => item?.ownerUid === tempShelfUserId);
      }
      loadedTemps.sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      const updateTempItems = snapshot.metadata.hasPendingWrites && !isUndoApplyingRef.current
        ? setTempItems
        : syncTempItems;
      updateTempItems((prev) => (isSameTransferItemList(prev, loadedTemps) ? prev : loadedTemps));
    }, (err) => {
      console.error("Temp Shelf Sync Error", err);
      const code = getFirestoreErrorCode(err);
      if (code === 'permission-denied' && !useLegacyTempShelf) {
        console.warn("Switching temp shelf to legacy path due permission-denied on user path.");
        setUseLegacyTempShelf(true);
      }
    });

    return () => {
      unsubscribeTemp();
    };
  }, [
    isAuthenticated,
    setTempItems,
    syncTempItems,
    tempShelfCollection,
    tempShelfSyncSource,
    tempShelfUserId,
    useLegacyTempShelf
  ]);

  const requestConfirm = (message, action) => {
    setConfirmDialog({
      isOpen: true,
      message,
      onConfirm: async () => {
        await action();
        setConfirmDialog({ isOpen: false, message: '', onConfirm: null });
      }
    });
  };

  const showAlert = (message, title = "通知", closeOnBackdrop = false) => {
    setAlertDialog({ isOpen: true, message, title, closeOnBackdrop });
  };

  const showUndoNotice = useCallback((message, tone = 'success') => {
    if (undoNoticeTimerRef.current) clearTimeout(undoNoticeTimerRef.current);
    setUndoNotice({ message, tone });
    undoNoticeTimerRef.current = setTimeout(() => {
      setUndoNotice(null);
      undoNoticeTimerRef.current = null;
    }, 2600);
  }, []);

  const cloudWriteQueuesRef = useRef(new Map());

  const runCloudWrite = useCallback((writer, options = {}) => {
    if (USE_LOCAL_STORAGE) return writer();
    const key = options.key || 'global';
    const queues = cloudWriteQueuesRef.current;
    const prev = queues.get(key) || Promise.resolve();

    const chainedWrite = prev
      .catch(() => undefined)
      .then(() => retryAsync(() => writer(), { retries: 8, baseDelayMs: 80 }));

    const queued = chainedWrite.catch(() => undefined);
    queues.set(key, queued);

    return chainedWrite.finally(() => {
      if (queues.get(key) === queued) {
        queues.delete(key);
      }
    });
  }, []);

  const runCloudTransaction = useCallback((transactionWork, options = {}) => {
    return runCloudWrite(() => runTransaction(db, transactionWork), options);
  }, [runCloudWrite]);

  // direction: 'undo' は直前の操作を戻す、'redo' は戻した操作をやり直す。
  // redo スタックには undo 向きのまま entry を保持し、適用時のみ invertUndoEntry で反転する。
  const restoreTimelineEntry = useCallback(async (direction) => {
    const accountId = undoAccountId;
    const isUndo = direction === 'undo';
    const actionLabel = isUndo ? '戻す' : 'やり直す';
    const actionLabelStem = isUndo ? '戻し' : 'やり直し';
    if (!accountId || isUndoBusyRef.current) return;
    if (isLockedRef.current) {
      showUndoNotice(`ロック中は操作を${actionLabel}ことはできません。`, 'warning');
      return;
    }
    if (isProcessing || panelArrangeSession) {
      showUndoNotice(`処理またはホバリングを完了してから操作を${actionLabelStem}てください。`, 'warning');
      return;
    }

    const entry = isUndo ? getLatestUndoEntry(accountId) : getLatestRedoEntry(accountId);
    const removeEntry = isUndo ? removeUndoEntry : removeRedoEntry;
    if (!entry) {
      showUndoNotice(isUndo ? 'このアカウントで戻せる操作はありません。' : 'このアカウントでやり直せる操作はありません。', 'neutral');
      return;
    }
    const effectiveEntry = isUndo ? entry : invertUndoEntry(entry);

    if (undoEntryHasClientConflict(effectiveEntry, workspaceStateRef.current)) {
      removeEntry(accountId, entry.id);
      showUndoNotice(`別の更新が重なったため、安全のためこの操作は${actionLabelStem}ませんでした。`, 'warning');
      return;
    }

    if (countUndoEntryChanges(effectiveEntry) > 450) {
      removeEntry(accountId, entry.id);
      showUndoNotice(`一度に${actionLabel}データ量が大きいため、この操作は${actionLabel}ことができません。`, 'warning');
      return;
    }

    isUndoBusyRef.current = true;
    isUndoApplyingRef.current = true;
    try {
      if (!USE_LOCAL_STORAGE) {
        await restoreCloudUndoEntry({
          accountId,
          collections: {
            sheets: sheetsCollection,
            images: imagesCollection,
            tempItems: tempShelfCollection,
            excludedItems: excludedItemsCollection
          },
          entry: effectiveEntry,
          ownerUid: tempShelfUserId,
          runCloudTransaction,
          useLegacyTempShelf
        });
      }

      const restoredWorkspace = applyUndoEntryToWorkspace(effectiveEntry, workspaceStateRef.current);
      syncSheets(restoredWorkspace.sheets);
      syncImages(restoredWorkspace.images);
      syncTempItems(restoredWorkspace.tempItems);
      syncExcludedItems(restoredWorkspace.excludedItems);
      removeEntry(accountId, entry.id);
      // 適用済み entry を反対側のスタックへ移し、undo ⇔ redo を往復可能にする
      if (isUndo) {
        pushRedoEntry(accountId, entry);
      } else {
        pushUndoEntry(accountId, entry);
      }
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
      setIsLabelSelectionMode(false);
      showUndoNotice(isUndo ? '直前の編集操作を戻しました。' : '操作をやり直しました。');
    } catch (error) {
      console.error(`${direction} failed:`, error);
      if (error?.code === 'undo-conflict') {
        removeEntry(accountId, entry.id);
        showUndoNotice(`別のアカウントによる更新を検出したため、操作は${actionLabelStem}ませんでした。`, 'warning');
      } else {
        showUndoNotice(`操作を${actionLabel}ことができませんでした。通信状態を確認して再度お試しください。`, 'warning');
      }
    } finally {
      isUndoApplyingRef.current = false;
      isUndoBusyRef.current = false;
    }
  }, [
    excludedItemsCollection,
    getLatestRedoEntry,
    getLatestUndoEntry,
    imagesCollection,
    isProcessing,
    isUndoApplyingRef,
    isUndoBusyRef,
    panelArrangeSession,
    pushRedoEntry,
    pushUndoEntry,
    removeRedoEntry,
    removeUndoEntry,
    runCloudTransaction,
    sheetsCollection,
    showUndoNotice,
    syncExcludedItems,
    syncImages,
    syncSheets,
    syncTempItems,
    tempShelfCollection,
    tempShelfUserId,
    undoAccountId,
    useLegacyTempShelf,
    workspaceStateRef
  ]);

  const handleUndoLatest = useCallback(() => restoreTimelineEntry('undo'), [restoreTimelineEntry]);
  const handleRedoLatest = useCallback(() => restoreTimelineEntry('redo'), [restoreTimelineEntry]);

  useEffect(() => {
    const handleUndoShortcut = (event) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
      const key = String(event.key || '').toLowerCase();
      const isUndoKey = key === 'z' && !event.shiftKey;
      const isRedoKey = (key === 'z' && event.shiftKey) || (key === 'y' && !event.shiftKey);
      if (!isUndoKey && !isRedoKey) return;
      const target = event.target;
      const isEditable = target instanceof HTMLInputElement
        || target instanceof HTMLTextAreaElement
        || target instanceof HTMLSelectElement
        || target?.isContentEditable;
      if (isEditable) return;

      event.preventDefault();
      void (isUndoKey ? handleUndoLatest() : handleRedoLatest());
    };

    window.addEventListener('keydown', handleUndoShortcut);
    return () => window.removeEventListener('keydown', handleUndoShortcut);
  }, [handleRedoLatest, handleUndoLatest]);

  const flushWorkLogDelta = useCallback(async (delta) => {
    if (!delta?.user?.uid || !delta.dateKey) return;
    const documentId = `${delta.user.uid}_${delta.dateKey}`;

    if (USE_LOCAL_STORAGE) {
      const stored = JSON.parse(localStorage.getItem(LOCAL_WORK_LOGS_KEY) || '{}');
      const previous = stored[documentId] || {};
      stored[documentId] = {
        ...applyWorkLogDeltaToRecord(previous, delta),
        uid: delta.user.uid,
        email: delta.user.email || '',
        displayName: delta.user.displayName || 'ローカル利用者',
        dateKey: delta.dateKey,
        lastSeenAtMs: Date.now(),
        version: 1
      };
      localStorage.setItem(LOCAL_WORK_LOGS_KEY, JSON.stringify(stored));
      return;
    }

    if (!workLogsCollection) return;
    const actionStats = {};
    Object.entries(delta.actions || {}).forEach(([actionId, stats]) => {
      const action = WORK_ACTIONS[actionId] || WORK_ACTIONS.other;
      actionStats[actionId] = {
        label: action.label,
        count: increment(Math.max(0, Number(stats.count) || 0)),
        activeMs: increment(Math.max(0, Math.round(Number(stats.activeMs) || 0)))
      };
    });
    await runCloudWrite(() => setDoc(doc(workLogsCollection, documentId), {
      uid: delta.user.uid,
      email: delta.user.email || '',
      displayName: delta.user.displayName || delta.user.email || '不明なアカウント',
      dateKey: delta.dateKey,
      totalActiveMs: increment(Math.max(0, Math.round(Number(delta.totalActiveMs) || 0))),
      sessionCount: increment(Math.max(0, Number(delta.sessionCount) || 0)),
      actionStats,
      lastSeenAt: serverTimestamp(),
      version: 1
    }, { merge: true }), { key: `activity-log:${documentId}` });
  }, [runCloudWrite, workLogsCollection]);

  const workLogUser = useMemo(() => ({
    uid: firebaseUser?.uid || '',
    email: firebaseUser?.email || '',
    displayName: signedInUserName || firebaseUser?.displayName || firebaseUser?.email || (USE_LOCAL_STORAGE ? 'ローカル利用者' : '')
  }), [firebaseUser, signedInUserName]);
  const { flushWorkActivityNow } = useWorkActivityTracker({
    enabled: isAuthenticated && !!workLogUser.uid,
    user: workLogUser,
    onFlush: flushWorkLogDelta
  });

  const buildTempShelfPayload = useCallback((basePayload = {}) => {
    const payload = { ...basePayload };
    if (useLegacyTempShelf && tempShelfUserId) {
      payload.ownerUid = tempShelfUserId;
    }
    return payload;
  }, [useLegacyTempShelf, tempShelfUserId]);

  const migratedPanelsMapRef = useRef(new Set());
  useEffect(() => {
    if (USE_LOCAL_STORAGE || !isAuthenticated || !sheetsCollection) return;
    if (!Array.isArray(sheets) || sheets.length === 0) return;

    sheets.forEach((sheet) => {
      if (!sheet?.id) return;
      const hasPanelsMap = !!(sheet.panelsMap && typeof sheet.panelsMap === 'object');
      const hasLegacyPanels = sheet._hasLegacyPanels === true;
      if ((hasPanelsMap && !hasLegacyPanels) || migratedPanelsMapRef.current.has(sheet.id)) return;

      migratedPanelsMapRef.current.add(sheet.id);
      runCloudWrite(
        () => updateDoc(doc(sheetsCollection, sheet.id), {
          panelsMap: toPanelsMap(sheet.panels || buildDefaultPanels()),
          panels: deleteField()
        }),
        { key: `sheet:${sheet.id}` }
      ).catch((error) => {
        console.error("Panels map migration failed:", error);
        migratedPanelsMapRef.current.delete(sheet.id);
      });
    });
  }, [isAuthenticated, sheets, sheetsCollection, runCloudWrite]);

  // --- Sales CSV Import Logic ---
  const handleImportSalesCSV = async (file) => {
    if (isLockedRef.current) return;
    setIsProcessing(true);
    setProgressMessage("売上データを解析中...");
    try {
      const text = await readFileAutoEncoding(file);
      const normalizedText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const rows = normalizedText.split('\n');

      const salesMap = {};
      const startIndex = 2; // ヘッダー2行スキップ

      rows.slice(startIndex).forEach((row) => {
        if (!row.trim()) return;
        const cols = parseCSVLine(row);

        if (cols.length <= 17) return;

        const rawCode = cols[3];
        if (!rawCode) return;
        const code = normalizeCode(rawCode);

        const name = cols[1] || '';
        const spec = cols[2] || '';
        const countStr = cols[17].replace(/,/g, '');
        const count = parseInt(countStr) || 0;

        if (!salesMap[code]) {
          salesMap[code] = [];
        }
        salesMap[code].push({ name, spec, count });
      });

      if (USE_LOCAL_STORAGE) {
        try {
          await idbHelper.setItem('salesData', salesMap);
          setSalesData(salesMap);
          setProgressMessage("完了しました");
          setTimeout(() => {
            setIsProcessing(false);
            showAlert("売上データを取り込みました");
            setIsSettingsOpen(false);
          }, 500);
        } catch (e) {
          console.error("Failed to save sales data to IDB:", e);
          showAlert("売上データの保存に失敗しました。");
          setIsProcessing(false);
        }
        return;
      }

      // Chunking logic
      setProgressMessage("データを保存中...");
      const entries = Object.entries(salesMap);
      const CHUNK_SIZE = 1000;
      const chunks = [];

      for (let i = 0; i < entries.length; i += CHUNK_SIZE) {
        const chunkEntries = entries.slice(i, i + CHUNK_SIZE);
        const chunkData = Object.fromEntries(chunkEntries);
        chunks.push(chunkData);
      }

      const batch = writeBatch(db);

      const snapshot = await getDocs(salesChunksCollection);
      snapshot.docs.forEach(d => batch.delete(d.ref));

      chunks.forEach((chunk, index) => {
        const docRef = doc(salesChunksCollection, `chunk_${index}`);
        batch.set(docRef, {
          items: JSON.stringify(chunk),
          updatedAt: serverTimestamp(),
          chunkIndex: index
        });
      });

      if (snapshot.size + chunks.length > 450) {
        const deleteBatch = writeBatch(db);
        snapshot.docs.forEach(d => deleteBatch.delete(d.ref));
        await runCloudWrite(() => deleteBatch.commit(), { key: 'sales-data' });

        for (let i = 0; i < chunks.length; i += 400) {
          const writeBatchChunk = writeBatch(db);
          chunks.slice(i, i + 400).forEach((chunk, idx) => {
            const realIdx = i + idx;
            const docRef = doc(salesChunksCollection, `chunk_${realIdx}`);
            writeBatchChunk.set(docRef, {
              items: JSON.stringify(chunk),
              updatedAt: serverTimestamp()
            });
          });
          await runCloudWrite(() => writeBatchChunk.commit(), { key: 'sales-data' });
        }
      } else {
        await runCloudWrite(() => batch.commit(), { key: 'sales-data' });
      }

      await setDoc(doc(settingsCollection, 'salesDataMeta'), {
        updatedAt: serverTimestamp(),
        totalItems: entries.length
      });

      const cachedMetaSeconds = Math.floor(Date.now() / 1000);
      await idbHelper.setItem(CLOUD_SALES_CACHE_KEY, {
        data: salesMap,
        metaSeconds: cachedMetaSeconds,
        fetchedAt: Date.now()
      });
      setSalesData(salesMap);
      setSalesDataLastUpdated(new Date(cachedMetaSeconds * 1000));

      showAlert("売上データを取り込みました！");
      setIsSettingsOpen(false);

    } catch (err) {
      console.error(err);
      showAlert("取り込みに失敗しました: " + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // --- Sales Hover Handler (Improved) ---
  const handleHoverSales = useCallback((data, pos) => {
    if (closeTimeoutRef.current) {
      clearTimeout(closeTimeoutRef.current);
      closeTimeoutRef.current = null;
    }
    if (data) {
      setHoveredSalesData(data);
      if (pos) setSalesPopupPos(pos);
    }
  }, []);

  const handleLeaveSales = useCallback(() => {
    if (closeTimeoutRef.current) clearTimeout(closeTimeoutRef.current);
    closeTimeoutRef.current = setTimeout(() => {
      setHoveredSalesData(null);
      setSalesPopupPos(null);
    }, 300); // 300ms delay to allow moving to popup
  }, []);

  const clearSalesModeLongPressTimer = useCallback(() => {
    if (salesModeLongPressTimerRef.current) {
      clearTimeout(salesModeLongPressTimerRef.current);
      salesModeLongPressTimerRef.current = null;
    }
  }, []);

  const startSalesModeLongPress = useCallback((event) => {
    if (panelArrangeSession) return;
    if (event?.button !== undefined && event.button !== 0) return;
    clearSalesModeLongPressTimer();
    salesModeLongPressTriggeredRef.current = false;
    salesModeLongPressTimerRef.current = setTimeout(() => {
      salesModeLongPressTriggeredRef.current = true;
      setIsSalesLookupOpen(true);
    }, 2000);
  }, [clearSalesModeLongPressTimer, panelArrangeSession]);

  const endSalesModeLongPress = useCallback(() => {
    clearSalesModeLongPressTimer();
  }, [clearSalesModeLongPressTimer]);

  const handleSalesModeButtonClick = useCallback(() => {
    if (panelArrangeSession) return;
    if (salesModeLongPressTriggeredRef.current) {
      salesModeLongPressTriggeredRef.current = false;
      return;
    }
    setIsSalesMode((prev) => !prev);
  }, [panelArrangeSession]);

  useEffect(() => {
    return () => {
      clearSalesModeLongPressTimer();
    };
  }, [clearSalesModeLongPressTimer]);

  // --- Image Logic (Duplicate Check) ---
  const checkImageUsage = (imageSrc) => {
    if (!imageSrc) return false;
    for (const sheet of sheets) {
      for (const panel of sheet.panels) {
        if (panel.image === imageSrc) return true;
      }
    }
    return false;
  };

  // --- Selection Logic ---
  const toggleMergeMode = () => {
    const newMode = !isMergeMode;
    setIsMergeMode(newMode);
    if (!newMode) setSelection({ sheetId: null, indices: [] });
  };

  const handleSelectPanel = useCallback((sheetId, index) => {
    if (isLockedRef.current) return;
    setSelection(prev => {
      if (prev.sheetId !== sheetId) return { sheetId, indices: [index] };
      const alreadySelected = prev.indices.includes(index);
      const newIndices = alreadySelected
        ? prev.indices.filter(i => i !== index)
        : [...prev.indices, index];
      return { sheetId, indices: newIndices };
    });
  }, []);

  const canMerge = useMemo(() => {
    if (!selection.sheetId || selection.indices.length < 2) return false;
    const sheet = sheets.find(s => s.id === selection.sheetId);
    if (!sheet) return false;

    const validPanels = selection.indices.every(idx => {
      const p = sheet.panels[idx];
      return !p.hidden && (p.rowSpan || 1) === 1 && (p.colSpan || 1) === 1;
    });
    if (!validPanels) return false;

    const coords = selection.indices.map(getCoords);
    const minRow = Math.min(...coords.map(c => c.row));
    const maxRow = Math.max(...coords.map(c => c.row));
    const minCol = Math.min(...coords.map(c => c.col));
    const maxCol = Math.max(...coords.map(c => c.col));

    const count = (maxRow - minRow + 1) * (maxCol - minCol + 1);
    if (count !== selection.indices.length) return false;

    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const idx = r * 4 + c;
        if (!selection.indices.includes(idx)) return false;
      }
    }

    return true;
  }, [selection, sheets]);

  const canSplit = useMemo(() => {
    if (!selection.sheetId || selection.indices.length === 0) return false;
    const sheet = sheets.find(s => s.id === selection.sheetId);
    if (!sheet) return false;

    return selection.indices.some(idx => {
      const p = sheet.panels[idx];
      return (p.rowSpan || 1) > 1 || (p.colSpan || 1) > 1;
    });
  }, [selection, sheets]);

  const handleMerge = useCallback(async () => {
    if (isLockedRef.current) return;
    if (!canMerge) return;
    const { sheetId, indices } = selection;
    if (!sheetId || indices.length === 0) return;

    const coords = indices.map(getCoords);
    const minRow = Math.min(...coords.map(c => c.row));
    const maxRow = Math.max(...coords.map(c => c.row));
    const minCol = Math.min(...coords.map(c => c.col));
    const maxCol = Math.max(...coords.map(c => c.col));
    const rowSpan = maxRow - minRow + 1;
    const colSpan = maxCol - minCol + 1;
    const primaryIndex = minRow * 4 + minCol;
    const sizeType = getSizeType(rowSpan, colSpan);
    const isArrangingSheet = panelArrangeSession?.sheetId === sheetId;

    if (USE_LOCAL_STORAGE) {
      const sheet = sheets.find(s => s.id === sheetId);
      if (!sheet) return;
      const newPanels = [...sheet.panels];
      newPanels[primaryIndex] = { ...newPanels[primaryIndex], rowSpan, colSpan, hidden: false, sizeType };
      indices.forEach(idx => {
        if (idx !== primaryIndex) {
          newPanels[idx] = isArrangingSheet
            ? { ...newPanels[idx], hidden: true, rowSpan: 1, colSpan: 1 }
            : { ...newPanels[idx], hidden: true, image: null, text: '', rowSpan: 1, colSpan: 1 };
        }
      });

      const newSheets = sheets.map(s =>
        s.id === sheetId ? { ...s, panels: newPanels } : s
      );
      setSheets(newSheets);
      if (isArrangingSheet) {
        setPanelArrangeSession((previous) => reconcilePanelArrangeSession(previous, newPanels));
      }
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
      return;
    }

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      let arrangedPanelsAfterMerge = null;
      await runCloudTransaction(async (transaction) => {
        const snap = await transaction.get(sheetRef);
        if (!snap.exists()) return;

        const serverPanels = getPanelsFromDocData(snap.data() || {});
        const validPanels = indices.every(idx => {
          const panel = serverPanels[idx] || {};
          return !panel.hidden && (panel.rowSpan || 1) === 1 && (panel.colSpan || 1) === 1;
        });
        if (!validPanels) return;

        const nextPanels = [...serverPanels];
        nextPanels[primaryIndex] = { ...nextPanels[primaryIndex], rowSpan, colSpan, hidden: false, sizeType };
        indices.forEach(idx => {
          if (idx !== primaryIndex) {
            nextPanels[idx] = isArrangingSheet
              ? { ...nextPanels[idx], hidden: true, rowSpan: 1, colSpan: 1 }
              : { ...nextPanels[idx], hidden: true, image: null, text: '', rowSpan: 1, colSpan: 1 };
          }
        });
        arrangedPanelsAfterMerge = nextPanels;

        const panelUpdates = buildPanelMapUpdates(serverPanels, nextPanels);
        if (Object.keys(panelUpdates).length > 0) {
          transaction.update(sheetRef, panelUpdates);
        }
      }, { key: `sheet:${sheetId}` });
      if (isArrangingSheet && arrangedPanelsAfterMerge) {
        setPanelArrangeSession((previous) => reconcilePanelArrangeSession(previous, arrangedPanelsAfterMerge));
      }
    } catch (error) {
      console.error("Merge error:", error);
      showAlert(buildFirestoreActionErrorMessage("結合に失敗しました。しばらく待って再試行してください。", error));
    } finally {
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
    }
  }, [canMerge, panelArrangeSession, selection, setSheets, sheets, sheetsCollection, runCloudTransaction, showAlert]);

  const handleSplit = useCallback(async () => {
    if (isLockedRef.current) return;
    const { sheetId, indices } = selection;
    if (!sheetId || indices.length === 0) return;

    if (USE_LOCAL_STORAGE) {
      const sheet = sheets.find(s => s.id === sheetId);
      if (!sheet) return;

      const newPanels = [...sheet.panels];
      indices.forEach(idx => {
        const p = newPanels[idx];
        if ((p.rowSpan || 1) > 1 || (p.colSpan || 1) > 1) {
          const rSpan = p.rowSpan || 1;
          const cSpan = p.colSpan || 1;
          const startRow = Math.floor(idx / 4);
          const startCol = idx % 4;

          newPanels[idx] = { ...p, rowSpan: 1, colSpan: 1, sizeType: '1/16（1コマ）' };

          for (let r = 0; r < rSpan; r++) {
            for (let c = 0; c < cSpan; c++) {
              if (r === 0 && c === 0) continue;
              const tIdx = (startRow + r) * 4 + (startCol + c);
              if (tIdx < 16) {
                newPanels[tIdx] = { ...newPanels[tIdx], hidden: false, sizeType: '1/16（1コマ）' };
              }
            }
          }
        }
      });

      const newSheets = sheets.map(s =>
        s.id === sheetId ? { ...s, panels: newPanels } : s
      );
      setSheets(newSheets);
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
      return;
    }

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      await runCloudTransaction(async (transaction) => {
        const snap = await transaction.get(sheetRef);
        if (!snap.exists()) return;

        const serverPanels = getPanelsFromDocData(snap.data() || {});
        const nextPanels = [...serverPanels];

        indices.forEach(idx => {
          const p = nextPanels[idx] || {};
          if ((p.rowSpan || 1) > 1 || (p.colSpan || 1) > 1) {
            const rSpan = p.rowSpan || 1;
            const cSpan = p.colSpan || 1;
            const startRow = Math.floor(idx / 4);
            const startCol = idx % 4;

            nextPanels[idx] = { ...p, rowSpan: 1, colSpan: 1, sizeType: '1/16（1コマ）' };
            for (let r = 0; r < rSpan; r++) {
              for (let c = 0; c < cSpan; c++) {
                if (r === 0 && c === 0) continue;
                const tIdx = (startRow + r) * 4 + (startCol + c);
                if (tIdx < 16) {
                  nextPanels[tIdx] = { ...nextPanels[tIdx], hidden: false, sizeType: '1/16（1コマ）' };
                }
              }
            }
          }
        });

        const panelUpdates = buildPanelMapUpdates(serverPanels, nextPanels);
        if (Object.keys(panelUpdates).length > 0) {
          transaction.update(sheetRef, panelUpdates);
        }
      }, { key: `sheet:${sheetId}` });
    } catch (error) {
      console.error("Split error:", error);
      showAlert(buildFirestoreActionErrorMessage("分離に失敗しました。しばらく待って再試行してください。", error));
    } finally {
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
    }
  }, [selection, setSheets, sheets, sheetsCollection, runCloudTransaction, showAlert]);

  // --- Core Actions ---

  const handleAddSheet = useCallback(async () => {
    if (isLockedRef.current) return;
    // 認証チェック: LocalStorageモードならUI認証のみ、FirebaseモードならFirebase認証も確認
    if (!isAuthenticated) return;
    if (!USE_LOCAL_STORAGE && !auth.currentUser) return;

    const newOrder = sheets.length > 0 ? Math.max(...sheets.map(s => s.order || 0)) + 1 : 0;

    const defaultPanels = buildDefaultPanels();

    if (USE_LOCAL_STORAGE) {
      const newSheet = {
        id: idbHelper.generateId(),
        createdAt: { seconds: Date.now() / 1000 },
        genre: 'none',
        order: newOrder,
        panels: defaultPanels
      };

      const newSheets = [...sheets, newSheet];
      setSheets(newSheets);
      return;
    }

    try {
      await addDoc(sheetsCollection, {
        genre: 'none',
        order: newOrder,
        panelsMap: toPanelsMap(defaultPanels),
        createdAt: serverTimestamp()
      });
    } catch (e) {
      console.error("Error adding sheet: ", e);
      showAlert("ページの追加に失敗しました。");
    }
  }, [isAuthenticated, setSheets, sheets, sheetsCollection]);

  const handleUpdatePanel = useCallback(async (sheetId, panelIndex, newData) => {
    if (isLockedRef.current) return;
    const sheetToUpdate = sheets.find(s => s.id === sheetId);
    if (!sheetToUpdate) return;
    const currentPanel = sheetToUpdate.panels[panelIndex] || {};
    const panelPatch = getPanelDataPatch(currentPanel, newData);
    if (Object.keys(panelPatch).length === 0) return;

    if (USE_LOCAL_STORAGE) {
      const updatedPanels = [...sheetToUpdate.panels];
      updatedPanels[panelIndex] = { ...currentPanel, ...panelPatch };
      const newSheets = sheets.map(s =>
        s.id === sheetId ? { ...s, panels: updatedPanels } : s
      );
      setSheets(newSheets);
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      return;
    }

    if (!sheetsCollection) return;

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      const fieldUpdates = {};
      const cloudPatch = sanitizePanelData({
        ...currentPanel,
        ...panelPatch
      });
      Object.entries(panelPatch).forEach(([key, value]) => {
        const cloudValue = key === 'image' ? cloudPatch.image : value;
        fieldUpdates[`panelsMap.${panelIndex}.${key}`] = cloudValue === undefined ? null : cloudValue;
      });
      if (cloudPatch.imageId && currentPanel.image) {
        fieldUpdates[`panelsMap.${panelIndex}.image`] = null;
      }
      if (Object.keys(fieldUpdates).length === 0) return;
      await runCloudWrite(() => updateDoc(sheetRef, fieldUpdates), { key: `sheet:${sheetId}` });
    } catch (error) {
      console.error("Panel update transaction failed:", error);
      showAlert(buildFirestoreActionErrorMessage("コマの更新に失敗しました。少し待ってから再実行してください。", error));
    }
  }, [setSheets, sheets, sheetsCollection, showAlert, runCloudWrite]);

  // --- Temp & Excluded Logic (Restored) ---

  const handleMoveToTemp = async (sheetId, panelIndex, movedText) => {
    if (isLockedRef.current) return;
    if (USE_LOCAL_STORAGE) {
      const sheet = sheets.find(s => s.id === sheetId);
      if (!sheet) return;
      const panel = sheet.panels[panelIndex];
      if (!hasPanelTransferableContent(panel)) return;

      const newTempItem = {
        id: idbHelper.generateId(),
        ...getPanelTransferableContent(panel, movedText),
        originalName: "退避アイテム",
        createdAt: { seconds: Date.now() / 1000 }
      };

      const newTempItems = [newTempItem, ...tempItems];
      setTempItems(newTempItems);
      // localStorageHelper.setItem('tempItems', newTempItems); // Auto-save handles this

      const updatedPanels = [...sheet.panels];
      updatedPanels[panelIndex] = clearPanelTransferableContent(updatedPanels[panelIndex]);

      const newSheets = sheets.map(s => s.id === sheetId ? { ...s, panels: updatedPanels } : s);
      setSheets(newSheets);
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      return;
    }

    if (!sheetsCollection || !tempShelfCollection) return;

    const moveWithTransaction = async (targetTempCollection, forceLegacyOwner = false) => {
      const sheetRef = doc(sheetsCollection, sheetId);
      const tempRef = doc(targetTempCollection);

      await runCloudTransaction(async (transaction) => {
        const sheetSnap = await transaction.get(sheetRef);
        if (!sheetSnap.exists()) return;

        const serverPanels = getPanelsFromDocData(sheetSnap.data() || {});
        const sourcePanel = serverPanels[panelIndex] || {};
        if (!hasPanelTransferableContent(sourcePanel)) return;

        transaction.set(tempRef, {
          ...getPanelTransferableContent(sourcePanel, movedText),
          originalName: "退避アイテム",
          ...(forceLegacyOwner && tempShelfUserId ? { ownerUid: tempShelfUserId } : {}),
          createdAt: serverTimestamp()
        });

        const nextPanels = [...serverPanels];
        nextPanels[panelIndex] = clearPanelTransferableContent(sourcePanel);
        const panelUpdates = buildPanelMapUpdates(serverPanels, nextPanels);
        if (Object.keys(panelUpdates).length > 0) {
          transaction.update(sheetRef, panelUpdates);
        }
      }, { key: `sheet:${sheetId}` });
    };

    try {
      await moveWithTransaction(tempShelfCollection, useLegacyTempShelf);
    } catch (error) {
      let finalError = error;
      const code = getFirestoreErrorCode(error);
      const shouldFallbackToLegacy = (
        code === 'permission-denied'
        && !useLegacyTempShelf
        && !!legacyTempShelfCollection
        && !!tempShelfUserId
      );

      if (shouldFallbackToLegacy) {
        try {
          setUseLegacyTempShelf(true);
          await moveWithTransaction(legacyTempShelfCollection, true);
          return;
        } catch (fallbackError) {
          console.error("Move to temp fallback transaction failed:", fallbackError);
          finalError = fallbackError;
        }
      } else {
        console.error("Move to temp transaction failed:", error);
      }

      if (getFirestoreErrorCode(finalError) === 'permission-denied') {
        showAlert("仮置き場への移動権限がありません。管理者にFirestoreルールの反映をご依頼ください。");
      } else {
        showAlert(buildFirestoreActionErrorMessage("仮置き場への移動に失敗しました。少し待ってから再実行してください。", finalError));
      }
    }
  };

  const handleAddDragItemToTempShelf = useCallback(async (dragPayload = {}, resolvedAssignment = null) => {
    if (isLockedRef.current) return;
    const assignment = resolvedAssignment || extractPanelAssignmentFromDragPayload(dragPayload, '');
    if (!assignment) return false;

    if (assignment.fromTempId) {
      // 既に仮置き場にあるアイテムを仮置き場へドロップした場合は何もしない
      return true;
    }

    if (!hasPanelTransferableContent(assignment)) return false;

    const originalName = parseNullableDragValue(dragPayload.name) || '仮置きアイテム';
    const itemPayload = {
      ...getPanelTransferableContent(assignment),
      originalName
    };

    if (USE_LOCAL_STORAGE) {
      const newTempItem = {
        id: idbHelper.generateId(),
        ...itemPayload,
        createdAt: { seconds: Date.now() / 1000 }
      };
      setTempItems((prev) => [newTempItem, ...(prev || [])]);

      if (assignment.fromExcludedId) {
        setExcludedItems((prev) => (prev || []).filter((item) => item.id !== assignment.fromExcludedId));
      }
      return true;
    }

    if (!tempShelfCollection) return false;

    const addToCollection = async (targetTempCollection, forceLegacyOwner = false) => {
      const batch = writeBatch(db);
      const tempRef = doc(targetTempCollection);
      batch.set(tempRef, {
        ...itemPayload,
        ...(forceLegacyOwner && tempShelfUserId ? { ownerUid: tempShelfUserId } : buildTempShelfPayload({})),
        createdAt: serverTimestamp()
      });

      if (assignment.fromExcludedId && excludedItemsCollection) {
        batch.delete(doc(excludedItemsCollection, assignment.fromExcludedId));
      }

      await runCloudWrite(() => batch.commit(), { key: 'temp-shelf' });
    };

    try {
      await addToCollection(tempShelfCollection, useLegacyTempShelf);
      return true;
    } catch (error) {
      let finalError = error;
      const code = getFirestoreErrorCode(error);
      const shouldFallbackToLegacy = (
        code === 'permission-denied'
        && !useLegacyTempShelf
        && !!legacyTempShelfCollection
        && !!tempShelfUserId
      );

      if (shouldFallbackToLegacy) {
        try {
          setUseLegacyTempShelf(true);
          await addToCollection(legacyTempShelfCollection, true);
          return true;
        } catch (fallbackError) {
          console.error("Add drag item fallback failed:", fallbackError);
          finalError = fallbackError;
        }
      } else {
        console.error("Add drag item to temp shelf failed:", error);
      }

      if (getFirestoreErrorCode(finalError) === 'permission-denied') {
        showAlert("仮置き場への追加権限がありません。管理者にFirestoreルールの反映をご依頼ください。");
      } else {
        showAlert(buildFirestoreActionErrorMessage("仮置き場への追加に失敗しました。少し待ってから再実行してください。", finalError));
      }
      return false;
    }
  }, [tempShelfCollection, excludedItemsCollection, runCloudWrite, showAlert, useLegacyTempShelf, legacyTempShelfCollection, tempShelfUserId, buildTempShelfPayload, setExcludedItems, setTempItems]);

  const handleDeleteFromTemp = async (id) => {
    if (isLockedRef.current) return;
    if (USE_LOCAL_STORAGE) {
      const newTempItems = tempItems.filter(item => item.id !== id);
      setTempItems(newTempItems);
      // localStorageHelper.setItem('tempItems', newTempItems); // Auto-save handles this
      return;
    }
    try {
      await runCloudWrite(() => deleteDoc(doc(tempShelfCollection, id)), { key: 'temp-shelf' });
    } catch (error) {
      const code = getFirestoreErrorCode(error);
      const shouldFallbackToLegacy = (
        code === 'permission-denied'
        && !useLegacyTempShelf
        && !!legacyTempShelfCollection
      );
      if (shouldFallbackToLegacy) {
        try {
          setUseLegacyTempShelf(true);
          await runCloudWrite(() => deleteDoc(doc(legacyTempShelfCollection, id)), { key: 'temp-shelf' });
          return;
        } catch (fallbackError) {
          console.error("Delete temp item fallback failed:", fallbackError);
        }
      } else {
        console.error("Delete temp item failed:", error);
      }
      showAlert("仮置き場アイテムの削除に失敗しました。");
    }
  };

  const handleMoveToExcluded = async (sheetId, panelIndex, movedText) => {
    if (isLockedRef.current) return;
    const currentSheet = sheets.find(s => s.id === sheetId);
    const currentPanel = currentSheet?.panels?.[panelIndex];
    if (!currentPanel || !hasPanelTransferableContent(currentPanel)) return;

    if (USE_LOCAL_STORAGE) {
      const newExcludedItem = {
        id: idbHelper.generateId(),
        ...getPanelTransferableContent(currentPanel, movedText),
        originalName: "掲載除外",
        createdAt: { seconds: Date.now() / 1000 }
      };

      const newExcludedItems = [newExcludedItem, ...excludedItems];
      setExcludedItems(newExcludedItems);
      // localStorageHelper.setItem('excludedItems', newExcludedItems); // Auto-save handles this

      const updatedPanels = [...currentSheet.panels];
      updatedPanels[panelIndex] = clearPanelTransferableContent(updatedPanels[panelIndex]);

      const newSheets = sheets.map(s => s.id === sheetId ? { ...currentSheet, panels: updatedPanels } : s);
      setSheets(newSheets);
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      return;
    }

    if (!sheetsCollection || !excludedItemsCollection) return;

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      const excludedRef = doc(excludedItemsCollection);

      await runCloudTransaction(async (transaction) => {
        const sheetSnap = await transaction.get(sheetRef);
        if (!sheetSnap.exists()) return;

        const serverPanels = getPanelsFromDocData(sheetSnap.data() || {});
        const sourcePanel = serverPanels[panelIndex] || {};
        if (!hasPanelTransferableContent(sourcePanel)) return;

        transaction.set(excludedRef, {
          ...getPanelTransferableContent(currentPanel, movedText),
          originalName: "掲載除外",
          createdAt: serverTimestamp()
        });

        const nextPanels = [...serverPanels];
        nextPanels[panelIndex] = clearPanelTransferableContent(sourcePanel);
        const panelUpdates = buildPanelMapUpdates(serverPanels, nextPanels);
        if (Object.keys(panelUpdates).length > 0) {
          transaction.update(sheetRef, panelUpdates);
        }
      }, { key: `sheet:${sheetId}` });
    } catch (error) {
      console.error("Move to excluded transaction failed:", error);
      showAlert(buildFirestoreActionErrorMessage("除外リストへの移動に失敗しました。少し待ってから再実行してください。", error));
    }
  };

  const handleDeleteFromExcluded = async (id) => {
    if (isLockedRef.current) return;
    requestConfirm(
      "掲載除外リストから完全に削除しますか？\n（復元できません）",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newExcludedItems = excludedItems.filter(item => item.id !== id);
          setExcludedItems(newExcludedItems);
          // localStorageHelper.setItem('excludedItems', newExcludedItems); // Auto-save handles this
          return;
        }
        try {
          await runCloudWrite(() => deleteDoc(doc(excludedItemsCollection, id)), { key: 'excluded-items' });
        } catch (error) {
          console.error("Delete excluded item failed:", error);
          showAlert("除外アイテムの削除に失敗しました。");
        }
      }
    );
  };

  // 除外リスト一括削除機能
  const handleBulkDeleteExcluded = async () => {
    if (isLockedRef.current) return;
    if (excludedItems.length === 0) return;
    requestConfirm(
      `除外リスト内の ${excludedItems.length} 件のアイテムを全て削除しますか？\n（復元できません）`,
      async () => {
        if (USE_LOCAL_STORAGE) {
          setExcludedItems([]);
          // localStorageHelper.setItem('excludedItems', []); // Auto-save handles this
          showAlert("除外リストを空にしました。");
          return;
        }

        const batch = writeBatch(db);
        excludedItems.forEach(item => {
          batch.delete(doc(excludedItemsCollection, item.id));
        });
        try {
          await runCloudWrite(() => batch.commit(), { key: 'excluded-items' });
          showAlert("除外リストを空にしました。");
        } catch (err) {
          console.error("Bulk delete excluded failed", err);
          showAlert("一括削除に失敗しました。");
        }
      }
    );
  };

  const handleExportExcludedCSV = () => {
    const csvContent = buildExcludedItemsCsvContent(excludedItems);
    downloadTextFile(csvContent, buildDatedCsvFilename('excluded_items'));
  };

  const removeMatchingTempItemsForImage = useCallback((assignedImage, assignedImageId) => {
    if (!assignedImage && !assignedImageId) return;

    const matchedItems = tempItems.filter((item) => {
      if (!item?.id) return false;
      if (getPanelFreeLabels(item).length > 0) return false;
      const byId = !!assignedImageId && !!item.imageId && item.imageId === assignedImageId;
      const byData = !!assignedImage && !!item.image && item.image === assignedImage;
      return byId || byData;
    });

    if (matchedItems.length === 0) return;

    if (USE_LOCAL_STORAGE) {
      const matchedIds = new Set(matchedItems.map((item) => item.id));
      setTempItems((prev) => prev.filter((item) => !matchedIds.has(item.id)));
      return;
    }

    if (!tempShelfCollection) return;

    const deleteBatch = writeBatch(db);
    matchedItems.forEach((item) => {
      deleteBatch.delete(doc(tempShelfCollection, item.id));
    });

    runCloudWrite(() => deleteBatch.commit(), { key: 'temp-shelf' }).catch((error) => {
      const code = getFirestoreErrorCode(error);
      if (code === 'permission-denied' && !useLegacyTempShelf) {
        setUseLegacyTempShelf(true);
        return;
      }
      console.error("Delete matched temp items failed:", error);
    });
  }, [setTempItems, tempItems, tempShelfCollection, runCloudWrite, useLegacyTempShelf]);

  const handlePanelUpdateWithCheck = (sheetId, panelIndex, newData) => {
    if (isLockedRef.current) return;
    const sanitizedData = { ...newData };
    const cameFromTemp = !!sanitizedData.fromTempId;
    const cameFromExcluded = !!sanitizedData.fromExcludedId;

    if (sanitizedData.fromTempId) {
      handleDeleteFromTemp(sanitizedData.fromTempId);
      delete sanitizedData.fromTempId;
    }
    if (sanitizedData.fromExcludedId) {
      if (USE_LOCAL_STORAGE) {
        const newExcludedItems = excludedItems.filter(item => item.id !== sanitizedData.fromExcludedId);
        setExcludedItems(newExcludedItems);
        // localStorageHelper.setItem('excludedItems', newExcludedItems); // Auto-saveに任せる
      } else {
        runCloudWrite(() => deleteDoc(doc(excludedItemsCollection, sanitizedData.fromExcludedId)), { key: 'excluded-items' }).catch((error) => {
          console.error("Delete excluded item during drop failed:", error);
        });
      }
      delete sanitizedData.fromExcludedId;
    }

    const isLibraryImageDrop = !cameFromTemp
      && !cameFromExcluded
      && !!sanitizedData.image
      && !sanitizedData.label
      && !sanitizedData.isText;
    if (isLibraryImageDrop) {
      removeMatchingTempItemsForImage(sanitizedData.image, sanitizedData.imageId);
    }

    const currentSheet = sheets.find(s => s.id === sheetId);
    const currentPanel = currentSheet?.panels[panelIndex];

    if (sanitizedData.image && sanitizedData.image !== currentPanel?.image && !sanitizedData.label) {
      if (checkImageUsage(sanitizedData.image)) {
        requestConfirm("同じ画像が既にはめ込まれています。\n配置しますか？", () => handleUpdatePanel(sheetId, panelIndex, sanitizedData));
        return;
      }
    }
    handleUpdatePanel(sheetId, panelIndex, sanitizedData);
  };

  const handleMoveToStock = async (sheetId, panelIndex, movedText) => {
    if (isLockedRef.current) return;
    const sheet = sheets.find(s => s.id === sheetId);
    if (!sheet) return;
    const panel = sheet.panels[panelIndex];
    if (!panel) return;

    const stockImage = images.find((image) => (
      (!!panel.imageId && image?.id === panel.imageId)
      || (!!panel.image && image?.data === panel.image)
    ));
    const panelImage = panel.image || stockImage?.data || null;

    // ダミー/テキスト/画像を持たないラベルは、画像ライブラリでは表現できないため仮置き場へ退避する。
    if (panel.label || panel.isText || !panelImage) {
      await handleMoveToTemp(sheetId, panelIndex, movedText);
      return;
    }

    const libraryMetadata = {
      code: panel.code || stockImage?.code || null,
      freeLabels: getPanelFreeLabels(panel),
      freeText: null
    };
    // コマから外す操作も「自分が作業した画像」として記録する
    const mergedWorkedBy = undoAccountId
      ? Array.from(new Set([...(stockImage?.workedBy || []), undoAccountId]))
      : (stockImage?.workedBy || null);

    try {
      let returnedImage;
      if (stockImage) {
        returnedImage = { ...stockImage, ...libraryMetadata, workedBy: mergedWorkedBy };
        if (!USE_LOCAL_STORAGE) {
          if (!imagesCollection) return;
          await runCloudWrite(
            () => updateDoc(doc(imagesCollection, stockImage.id), {
              ...libraryMetadata,
              ...(undoAccountId ? { workedBy: arrayUnion(undoAccountId) } : {})
            }),
            { key: 'images' }
          );
        }
      } else {
        const fallbackName = panel.code ? `${panel.code}.png` : `returned-${Date.now()}.png`;
        returnedImage = {
          id: idbHelper.generateId(),
          name: fallbackName,
          data: panelImage,
          ...libraryMetadata,
          workedBy: undoAccountId ? [undoAccountId] : null,
          createdAt: { seconds: Date.now() / 1000 }
        };

        if (!USE_LOCAL_STORAGE) {
          if (!imagesCollection) return;
          const imageRef = await addDoc(imagesCollection, {
            name: fallbackName,
            data: panelImage,
            ...libraryMetadata,
            ...(undoAccountId ? { workedBy: [undoAccountId] } : {}),
            createdAt: serverTimestamp()
          });
          returnedImage = { ...returnedImage, id: imageRef.id };
        }
      }

      const nextImages = normalizeStockImages(
        stockImage
          ? images.map((image) => image.id === stockImage.id ? returnedImage : image)
          : [...images, returnedImage]
      );
      setImages(nextImages);
      persistCloudImagesCache(nextImages);
      handlePanelUpdateWithCheck(sheetId, panelIndex, clearPanelTransferableContent(panel));
    } catch (error) {
      console.error("Move to stock failed", error);
      showAlert(buildFirestoreActionErrorMessage("画像ライブラリへの移動に失敗しました。元のコマは保持されています。", error));
    }
  };

  const handleMovePanel = async (fromSheetId, fromIndex, toSheetId, toIndex, movedText) => {
    if (isLockedRef.current) return;
    if (fromSheetId === toSheetId && fromIndex === toIndex) return;

    if (USE_LOCAL_STORAGE) {
      const fromSheet = sheets.find(s => s.id === fromSheetId);
      const toSheet = sheets.find(s => s.id === toSheetId);
      if (!fromSheet || !toSheet) return;

      const fromPanel = fromSheet.panels[fromIndex];
      if (!hasPanelTransferableContent(fromPanel)) return;

      if (fromSheetId === toSheetId) {
        const newPanels = [...fromSheet.panels];
        const dataToMove = { ...newPanels[fromIndex] };

        newPanels[fromIndex] = clearPanelTransferableContent(newPanels[fromIndex]);
        newPanels[toIndex] = applyPanelTransferableContent(newPanels[toIndex], dataToMove, movedText);

        const newSheets = sheets.map(s => s.id === fromSheetId ? { ...s, panels: newPanels } : s);
        setSheets(newSheets);
        // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      } else {
        const newFromPanels = [...fromSheet.panels];
        const dataToMove = { ...newFromPanels[fromIndex] };

        newFromPanels[fromIndex] = clearPanelTransferableContent(newFromPanels[fromIndex]);

        const newToPanels = [...toSheet.panels];
        newToPanels[toIndex] = applyPanelTransferableContent(newToPanels[toIndex], dataToMove, movedText);

        const newSheets = sheets.map(s => {
          if (s.id === fromSheetId) return { ...s, panels: newFromPanels };
          if (s.id === toSheetId) return { ...s, panels: newToPanels };
          return s;
        });
        setSheets(newSheets);
        // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      }
      return;
    }

    try {
      const fromSheetRef = doc(sheetsCollection, fromSheetId);
      const toSheetRef = doc(sheetsCollection, toSheetId);

      await runCloudTransaction(async (transaction) => {
        const fromSnap = await transaction.get(fromSheetRef);
        if (!fromSnap.exists()) return;

        const fromPanels = getPanelsFromDocData(fromSnap.data() || {});
        const dataToMove = { ...(fromPanels[fromIndex] || {}) };
        if (!hasPanelTransferableContent(dataToMove)) return;

        if (fromSheetId === toSheetId) {
          const nextPanels = [...fromPanels];
          const targetPanel = nextPanels[toIndex] || {};
          nextPanels[fromIndex] = clearPanelTransferableContent(dataToMove);
          nextPanels[toIndex] = applyPanelTransferableContent(targetPanel, dataToMove, movedText);
          const panelUpdates = buildPanelMapUpdates(fromPanels, nextPanels);
          if (Object.keys(panelUpdates).length > 0) {
            transaction.update(fromSheetRef, panelUpdates);
          }
          return;
        }

        const toSnap = await transaction.get(toSheetRef);
        if (!toSnap.exists()) return;

        const toPanels = getPanelsFromDocData(toSnap.data() || {});
        const nextFromPanels = [...fromPanels];
        const nextToPanels = [...toPanels];
        nextFromPanels[fromIndex] = clearPanelTransferableContent(dataToMove);
        const targetPanel = nextToPanels[toIndex] || {};
        nextToPanels[toIndex] = applyPanelTransferableContent(targetPanel, dataToMove, movedText);

        const fromUpdates = buildPanelMapUpdates(fromPanels, nextFromPanels);
        const toUpdates = buildPanelMapUpdates(toPanels, nextToPanels);
        if (Object.keys(fromUpdates).length > 0) {
          transaction.update(fromSheetRef, fromUpdates);
        }
        if (Object.keys(toUpdates).length > 0) {
          transaction.update(toSheetRef, toUpdates);
        }
      }, { key: `sheet-pair:${[fromSheetId, toSheetId].sort().join('|')}` });
    } catch (error) {
      console.error("Move panel transaction failed:", error);
      showAlert(buildFirestoreActionErrorMessage("コマの移動に失敗しました。少し待ってから再実行してください。", error));
    }
  };

  const applyDragPayloadToPanel = useCallback((targetSheetId, targetIndex, dragPayload = {}) => {
    if (!targetSheetId || Number.isNaN(targetIndex)) return false;

    const arrangePayload = extractPanelArrangeDragPayload(dragPayload);
    if (arrangePayload) {
      if (
        !panelArrangeSession
        || arrangePayload.sheetId !== panelArrangeSession.sheetId
        || targetSheetId !== panelArrangeSession.sheetId
      ) return false;

      const targetSheet = sheets.find((sheet) => sheet.id === targetSheetId);
      if (!targetSheet?.panels) return false;
      const staged = stagePanelArrangeDrop(
        panelArrangeSession,
        arrangePayload.tokenId,
        targetIndex,
        targetSheet.panels
      );
      if (staged.status === 'blocked-content') {
        showAlert('ダミーまたはテキストがあるコマには画像を重ねられません。空きコマを指定してください。');
        return false;
      }
      if (staged.status !== 'placed') return false;
      setPanelArrangeSession(staged.session);
      return true;
    }

    if (panelArrangeSession) return false;

    const movePayload = extractPanelMoveDragPayload(dragPayload) || getActivePanelMoveDragPayload();
    if (movePayload) {
      void handleMovePanel(
        movePayload.sourceSheetId,
        movePayload.sourceIndex,
        targetSheetId,
        targetIndex,
        movePayload.movedText
      );
      return true;
    }

    const targetSheet = sheets.find((sheet) => sheet.id === targetSheetId);
    const targetPanel = targetSheet?.panels?.[targetIndex];
    if (!targetPanel) return false;

    const assignment = extractPanelAssignmentFromDragPayload(
      dragPayload,
      targetPanel.text || 'テキストを入力'
    );
    if (!assignment) return false;

    handlePanelUpdateWithCheck(targetSheetId, targetIndex, {
      ...targetPanel,
      ...assignment
    });
    return true;
  }, [handleMovePanel, handlePanelUpdateWithCheck, panelArrangeSession, sheets, showAlert]);

  const applyDragPayloadToTempShelf = useCallback((dragPayload = {}) => {
    if (extractPanelArrangeDragPayload(dragPayload)) return false;
    const movePayload = extractPanelMoveDragPayload(dragPayload) || getActivePanelMoveDragPayload();
    if (movePayload) {
      void handleMoveToTemp(
        movePayload.sourceSheetId,
        movePayload.sourceIndex,
        movePayload.movedText
      );
      return true;
    }

    const assignment = extractPanelAssignmentFromDragPayload(dragPayload, '');
    if (!assignment) return false;
    if (assignment.fromTempId) return true;

    void handleAddDragItemToTempShelf(dragPayload, assignment);
    return true;
  }, [handleMoveToTemp, handleAddDragItemToTempShelf]);

  const applyDragPayloadToStockList = useCallback((dragPayload = {}) => {
    if (extractPanelArrangeDragPayload(dragPayload)) return false;
    const movePayload = extractPanelMoveDragPayload(dragPayload) || getActivePanelMoveDragPayload();
    if (!movePayload) return false;
    void handleMoveToStock(
      movePayload.sourceSheetId,
      movePayload.sourceIndex,
      movePayload.movedText
    );
    return true;
  }, [handleMoveToStock]);

  const applyDragPayloadToExcludedList = useCallback((dragPayload = {}) => {
    if (extractPanelArrangeDragPayload(dragPayload)) return false;
    const movePayload = extractPanelMoveDragPayload(dragPayload) || getActivePanelMoveDragPayload();
    if (!movePayload) return false;
    void handleMoveToExcluded(
      movePayload.sourceSheetId,
      movePayload.sourceIndex,
      movePayload.movedText
    );
    return true;
  }, [handleMoveToExcluded]);

  const finalizePanelArrangeMode = useCallback(async () => {
    if (!panelArrangeSession || isPanelArrangeFinalizing) return;
    if (isLockedRef.current) {
      showAlert('画面ロックを解除してからホバリングを解除してください。');
      return;
    }

    const currentSheet = sheets.find((sheet) => sheet.id === panelArrangeSession.sheetId);
    if (!currentSheet?.panels) {
      showAlert('対象ページを確認できないため、ホバリングを解除できません。');
      return;
    }

    const reconciled = reconcilePanelArrangeSession(panelArrangeSession, currentSheet.panels);
    const unresolvedCount = getUnresolvedPanelArrangeTokens(reconciled).length;
    if (unresolvedCount > 0) {
      setPanelArrangeSession(reconciled);
      showAlert(`浮遊した画像が ${unresolvedCount} 件残っています。すべての画像をコマへ割り付けてください。`);
      return;
    }

    setIsPanelArrangeFinalizing(true);
    try {
      if (USE_LOCAL_STORAGE) {
        const finalPanels = buildPanelArrangeFinalPanels(currentSheet.panels, reconciled);
        if (!finalPanels) throw new Error('panel-arrange-incomplete');
        setSheets((previous) => previous.map((sheet) => (
          sheet.id === reconciled.sheetId ? { ...sheet, panels: finalPanels } : sheet
        )));
        clearPanelArrangeModeState();
        return;
      }

      if (!sheetsCollection) return;
      let transactionUnresolvedCount = 0;
      await runCloudTransaction(async (transaction) => {
        const sheetRef = doc(sheetsCollection, reconciled.sheetId);
        const snapshot = await transaction.get(sheetRef);
        if (!snapshot.exists()) throw new Error('panel-arrange-sheet-missing');

        const serverPanels = getPanelsFromDocData(snapshot.data() || {});
        const serverSession = reconcilePanelArrangeSession(reconciled, serverPanels);
        transactionUnresolvedCount = getUnresolvedPanelArrangeTokens(serverSession).length;
        if (transactionUnresolvedCount > 0) return;

        const finalPanels = buildPanelArrangeFinalPanels(serverPanels, serverSession);
        if (!finalPanels) {
          transactionUnresolvedCount = 1;
          return;
        }
        const panelUpdates = buildPanelMapUpdates(serverPanels, finalPanels);
        if (Object.keys(panelUpdates).length > 0) {
          transaction.update(sheetRef, panelUpdates);
        }
      }, { key: `sheet:${reconciled.sheetId}` });

      if (transactionUnresolvedCount > 0) {
        showAlert(`ページ構成の変更により浮遊画像が ${transactionUnresolvedCount} 件残っています。割り付けを完了してください。`);
        return;
      }
      clearPanelArrangeModeState();
    } catch (error) {
      console.error('Panel arrange finalize failed:', error);
      showAlert(buildFirestoreActionErrorMessage('ホバリング中の画像配置を保存できませんでした。配置状態は保持されています。', error));
    } finally {
      setIsPanelArrangeFinalizing(false);
    }
  }, [
    clearPanelArrangeModeState,
    isPanelArrangeFinalizing,
    panelArrangeSession,
    runCloudTransaction,
    setSheets,
    sheets,
    sheetsCollection,
    showAlert
  ]);

  const dispatchPointerDropToZone = useCallback((zoneId, dragPayload = {}) => {
    if (!zoneId) return false;
    if (extractPanelArrangeDragPayload(dragPayload) && !zoneId.startsWith(DAIWARI_PANEL_DROPZONE_PREFIX)) {
      return false;
    }
    if (zoneId === 'temp') return applyDragPayloadToTempShelf(dragPayload);
    if (zoneId === 'stock') return applyDragPayloadToStockList(dragPayload);
    if (zoneId === 'excluded') return applyDragPayloadToExcludedList(dragPayload);
    if (!zoneId.startsWith(DAIWARI_PANEL_DROPZONE_PREFIX)) return false;

    const panelTarget = zoneId.slice(DAIWARI_PANEL_DROPZONE_PREFIX.length);
    const separatorIndex = panelTarget.lastIndexOf(':');
    if (separatorIndex === -1) return false;

    const sheetId = panelTarget.slice(0, separatorIndex);
    const panelIndex = Number.parseInt(panelTarget.slice(separatorIndex + 1), 10);
    if (!sheetId || Number.isNaN(panelIndex)) return false;

    return applyDragPayloadToPanel(sheetId, panelIndex, dragPayload);
  }, [applyDragPayloadToTempShelf, applyDragPayloadToStockList, applyDragPayloadToExcludedList, applyDragPayloadToPanel]);

  const handlePointerDropAtPoint = useCallback((clientX, clientY, dragPayload = {}) => {
    if (typeof document === 'undefined') return false;
    const target = document.elementFromPoint(clientX, clientY);
    const dropZoneEl = target?.closest?.(`[${DAIWARI_DROPZONE_ATTR}]`);
    const zoneId = dropZoneEl?.getAttribute?.(DAIWARI_DROPZONE_ATTR) || '';
    return dispatchPointerDropToZone(zoneId, dragPayload);
  }, [dispatchPointerDropToZone]);

  useEffect(() => {
    if (typeof document === 'undefined') return undefined;

    const handleDocumentDragOver = (event) => {
      if (!getActiveNativeDragPayload()) return;
      event.preventDefault();
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = 'move';
      }
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
      const dropZoneEl = targetElement?.closest?.(`[${DAIWARI_DROPZONE_ATTR}]`);
      const zoneId = dropZoneEl?.getAttribute?.(DAIWARI_DROPZONE_ATTR) || '';
      if (!zoneId) {
        clearActiveNativeDragPayload();
        return;
      }

      const handled = dispatchPointerDropToZone(zoneId, dragPayload);
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
  }, [dispatchPointerDropToZone]);

  const positionPointerDragOverlay = useCallback((clientX, clientY) => {
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
          handled = handlePointerDropAtPoint(pointerEvent.clientX, pointerEvent.clientY, session.payload);
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
        requestAnimationFrame(() => positionPointerDragOverlay(session.currentX, session.currentY));
      }

      moveEvent.preventDefault();
      positionPointerDragOverlay(session.currentX, session.currentY);
    };

    session.handleUp = (upEvent) => {
      finishPointerDrag(upEvent, true);
    };

    session.handleCancel = (cancelEvent) => {
      finishPointerDrag(cancelEvent, false);
    };

    pointerDragSessionRef.current = session;

    try {
      event.currentTarget?.setPointerCapture?.(event.pointerId);
    } catch (error) {
      void error;
    }

    document.addEventListener('pointermove', session.handleMove, { passive: false });
    document.addEventListener('pointerup', session.handleUp, { passive: false });
    document.addEventListener('pointercancel', session.handleCancel, { passive: false });
  }, [clearPointerDragSession, handlePointerDropAtPoint, positionPointerDragOverlay]);

  useEffect(() => {
    const handleCaptureClick = (event) => {
      if (!suppressNextClickRef.current) return;
      suppressNextClickRef.current = false;
      event.preventDefault();
      event.stopPropagation();
    };

    document.addEventListener('click', handleCaptureClick, true);
    return () => {
      document.removeEventListener('click', handleCaptureClick, true);
    };
  }, []);

  useEffect(() => {
    return () => {
      clearPointerDragSession();
    };
  }, [clearPointerDragSession]);

  // --- Image & Bulk Actions ---

  const persistCloudImagesCache = useCallback((nextImages) => {
    if (USE_LOCAL_STORAGE) return;
    idbHelper.setItem(CLOUD_IMAGES_CACHE_KEY, {
      items: normalizeStockImages(nextImages || []),
      fetchedAt: Date.now()
    }).catch((error) => {
      console.error("Cloud image cache save failed:", error);
    });
  }, []);

  const handleUploadImage = async (e) => {
    if (isLockedRef.current) return;
    // 認証チェックを緩和（userオブジェクトではなくフラグで判定）
    if (!e.target.files || e.target.files.length === 0 || !isAuthenticated) return;
    const files = Array.from(e.target.files);

    let successCount = 0;
    let failCount = 0;
    const newImages = [];

    const uploadPromises = files.map(async (file) => {
      try {
        const compressedDataUrl = await compressImage(file);
        const newImage = {
          id: idbHelper.generateId(),
          name: file.name,
          data: compressedDataUrl,
          workedBy: undoAccountId ? [undoAccountId] : null,
          createdAt: { seconds: Date.now() / 1000 }
        };

        if (USE_LOCAL_STORAGE) {
          newImages.push(newImage);
        } else {
          const imageDocRef = await addDoc(imagesCollection, {
            name: file.name,
            data: compressedDataUrl,
            ...(undoAccountId ? { workedBy: [undoAccountId] } : {}),
            createdAt: serverTimestamp()
          });
          newImages.push({ ...newImage, id: imageDocRef.id });
        }
        successCount++;
      } catch (err) {
        console.error(`Failed to upload ${file.name}:`, err);
        failCount++;
      }
    });

    try {
      await Promise.all(uploadPromises);

      if (newImages.length > 0) {
        const updatedImages = normalizeStockImages([...images, ...newImages]);
        setImages(updatedImages);
        persistCloudImagesCache(updatedImages);
        // localStorageHelper.setItem('images', updatedImages); // Auto-save handles this
      }

      if (failCount > 0) {
        showAlert(`${successCount}枚の画像をアップロードしました。${failCount}枚は失敗しました。`);
      }
    } catch (err) {
      console.error("Batch upload error", err);
    }

    e.target.value = '';
  };

  const handleDeleteImage = (imgId, fallbackData = null) => {
    if (isLockedRef.current) return;
    requestConfirm(
      "画像をストックから削除しますか？",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newImages = images.filter(img => {
            if (imgId) return img.id !== imgId;
            if (fallbackData) return img.data !== fallbackData;
            return true;
          });
          setImages(newImages);
          // localStorageHelper.setItem('images', newImages); // Auto-save handles this
          return;
        }
        try {
          const deleteIdSet = new Set();
          if (imgId) deleteIdSet.add(imgId);

          if (fallbackData) {
            const snapshot = await getDocs(imagesCollection);
            snapshot.docs.forEach((imageDoc) => {
              const data = imageDoc.data() || {};
              const resolved = data.data || data.image || null;
              if (resolved && resolved === fallbackData) {
                deleteIdSet.add(imageDoc.id);
              }
            });
          }

          const deleteIds = Array.from(deleteIdSet).filter(Boolean);
          if (deleteIds.length === 0) {
            showAlert("削除対象の画像IDが見つかりませんでした。");
            return;
          }

          const batch = writeBatch(db);
          deleteIds.forEach((id) => {
            batch.delete(doc(imagesCollection, id));
          });
          await runCloudWrite(() => batch.commit(), { key: 'images' });
          const deleteIdLookup = new Set(deleteIds);
          const newImages = images.filter((img) => {
            if (img.id && deleteIdLookup.has(img.id)) return false;
            if (fallbackData && img.data === fallbackData) return false;
            return true;
          });
          setImages(newImages);
          persistCloudImagesCache(newImages);
        } catch (error) {
          console.error("Delete image failed:", error);
          showAlert("画像削除に失敗しました。");
        }
      }
    );
  };

  const handleBulkDeleteImages = async (imageTargets) => {
    if (isLockedRef.current) return;
    if (!imageTargets || imageTargets.length === 0) return;

    requestConfirm(
      `${imageTargets.length}枚の画像を削除しますか？`,
      async () => {
        const idSet = new Set();
        const dataSet = new Set();

        imageTargets.forEach((target) => {
          if (!target) return;
          if (typeof target === 'string') {
            if (target) idSet.add(target);
            return;
          }
          if (target.id) idSet.add(target.id);
          else if (target.data) dataSet.add(target.data);
        });

        if (USE_LOCAL_STORAGE) {
          const newImages = images.filter(img => {
            if (img.id) return !idSet.has(img.id);
            if (img.data) return !dataSet.has(img.data);
            return true;
          });
          setImages(newImages);
          // localStorageHelper.setItem('images', newImages); // Auto-save handles this
          return;
        }

        if (dataSet.size > 0) {
          const snapshot = await getDocs(imagesCollection);
          snapshot.docs.forEach((imageDoc) => {
            const data = imageDoc.data() || {};
            const resolved = data.data || data.image || null;
            if (resolved && dataSet.has(resolved)) {
              idSet.add(imageDoc.id);
            }
          });
        }

        const imageIds = Array.from(idSet).filter(Boolean);
        if (imageIds.length === 0) {
          showAlert("削除可能な画像IDが見つかりませんでした。");
          return;
        }

        try {
          const BATCH_LIMIT = 450;
          for (let i = 0; i < imageIds.length; i += BATCH_LIMIT) {
            const chunk = imageIds.slice(i, i + BATCH_LIMIT);
            const batch = writeBatch(db);
            chunk.forEach(id => {
              const ref = doc(imagesCollection, id);
              batch.delete(ref);
            });
            await runCloudWrite(() => batch.commit(), { key: 'images' });
          }
          const deleteIdLookup = new Set(imageIds);
          const newImages = images.filter((img) => {
            if (img.id && deleteIdLookup.has(img.id)) return false;
            if (img.data && dataSet.has(img.data)) return false;
            return true;
          });
          setImages(newImages);
          persistCloudImagesCache(newImages);
        } catch (err) {
          console.error("Bulk image delete failed", err);
          showAlert("一括削除に失敗しました");
        }
      }
    );
  };

  // --- Page Actions ---

  const handleNavigatePage = (direction) => {
    if (panelArrangeSession) {
      showAlert('ホバリングを解除してから前後のページへ移動してください。');
      return;
    }
    const navigationList = sheets;
    if (!navigationList || navigationList.length === 0) return;

    const foundIndex = navigationList.findIndex(s => s.id === activeSheetId);
    const currentIndex = foundIndex === -1 ? 0 : foundIndex;

    if (direction === 'prev' && currentIndex > 0) {
      setActiveSheetId(navigationList[currentIndex - 1].id);
      setIsLabelSelectionMode(false);
    } else if (direction === 'next' && currentIndex < navigationList.length - 1) {
      setActiveSheetId(navigationList[currentIndex + 1].id);
      setIsLabelSelectionMode(false);
    }
  };

  const handleChangeGenre = async (sheetId, newGenre) => {
    if (isLockedRef.current) return;
    if (USE_LOCAL_STORAGE) {
      const newSheets = sheets.map(s => s.id === sheetId ? { ...s, genre: newGenre } : s);
      setSheets(newSheets);
      return;
    }
    const sheetRef = doc(sheetsCollection, sheetId);
    try {
      await runCloudWrite(() => updateDoc(sheetRef, { genre: newGenre }), { key: `sheet:${sheetId}` });
    } catch (error) {
      console.error("Genre update failed:", error);
      showAlert("ジャンル変更に失敗しました。");
    }
  };

  const handleDeleteSheet = (sheetId) => {
    if (isLockedRef.current) return;
    requestConfirm(
      "このページを削除しますか？\n（この操作は取り消せません）",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newSheets = sheets.filter(s => s.id !== sheetId);
          setSheets(newSheets);
          if (activeSheetId === sheetId) setActiveSheetId(null);
          return;
        }
        try {
          await runCloudWrite(() => deleteDoc(doc(sheetsCollection, sheetId)), { key: `sheet:${sheetId}` });
        } catch (error) {
          console.error("Delete sheet failed:", error);
          showAlert("ページ削除に失敗しました。");
        }
      }
    );
  };

  const togglePageSelectionMode = () => {
    const newMode = !isPageSelectionMode;
    setIsPageSelectionMode(newMode);
    if (!newMode) {
      setSelectedSheetIds(new Set());
    } else {
      setViewMode('overview');
      setActiveSheetId(null);
      setIsLabelSelectionMode(false);
    }
  };

  const handleSelectAllPages = () => {
    if (selectedSheetIds.size === displaySheets.length) {
      setSelectedSheetIds(new Set());
    } else {
      setSelectedSheetIds(new Set(displaySheets.map(s => s.id)));
    }
  };

  const handleToggleSheetSelection = (sheetId) => {
    setSelectedSheetIds(prev => {
      const newSet = new Set(prev);
      if (newSet.has(sheetId)) {
        newSet.delete(sheetId);
      } else {
        newSet.add(sheetId);
      }
      return newSet;
    });
  };

  const handleExportSelectedPdf = async () => {
    if (isProcessing || selectedSheetIds.size === 0) return;
    const exportPlan = buildPdfExportPlan({ sheets, selectedSheetIds, genres: GENRES });
    if (exportPlan.pages.length === 0) {
      showAlert('PDFへ出力するページを選択してください。');
      return;
    }

    setIsProcessing(true);
    setProgressValue(0);
    setProgressMax(exportPlan.pages.length);
    setProgressMessage('PDF出力の準備をしています...');

    try {
      const pdfRenderer = await createPdfRenderer();
      for (let pageIndex = 0; pageIndex < exportPlan.pages.length; pageIndex++) {
        const page = exportPlan.pages[pageIndex];
        setPdfExportPage(page);
        setProgressMessage(`Page ${page.pageNumber}（${page.genreLabel}）をPDFへ変換しています...`);
        const surface = await waitForPdfExportSurface(page.sheet.id);
        await pdfRenderer.appendSurface(surface);
        setProgressValue(pageIndex + 1);
      }
      pdfRenderer.save(exportPlan.filename);
    } catch (error) {
      console.error('PDF export failed:', error);
      showAlert(`PDF出力に失敗しました。\n${error?.message || '時間をおいて再度お試しください。'}`);
    } finally {
      setPdfExportPage(null);
      setIsProcessing(false);
    }
  };

  const handleSwapPages = async () => {
    if (isLockedRef.current) return;
    if (selectedSheetIds.size !== 2) return;
    const [id1, id2] = Array.from(selectedSheetIds);
    const sheet1 = sheets.find(s => s.id === id1);
    const sheet2 = sheets.find(s => s.id === id2);

    if (!sheet1 || !sheet2) return;

    requestConfirm(
      "選択した2つのページの内容を入れ替えますか？",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newSheets = sheets.map(s => {
            if (s.id === id1) return { ...s, genre: sheet2.genre, panels: sheet2.panels };
            if (s.id === id2) return { ...s, genre: sheet1.genre, panels: sheet1.panels };
            return s;
          });
          setSheets(newSheets);
          // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
          return;
        }

        try {
          const ref1 = doc(sheetsCollection, id1);
          const ref2 = doc(sheetsCollection, id2);

          await runCloudTransaction(async (transaction) => {
            const snap1 = await transaction.get(ref1);
            const snap2 = await transaction.get(ref2);
            if (!snap1.exists() || !snap2.exists()) return;

            const data1 = snap1.data() || {};
            const data2 = snap2.data() || {};
            const data1Panels = getPanelsFromDocData(data1);
            const data2Panels = getPanelsFromDocData(data2);

            transaction.update(ref1, {
              genre: data2.genre || 'none',
              panelsMap: toPanelsMap(data2Panels)
            });
            transaction.update(ref2, {
              genre: data1.genre || 'none',
              panelsMap: toPanelsMap(data1Panels)
            });
          }, { key: `sheet-pair:${[id1, id2].sort().join('|')}` });

          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
        } catch (err) {
          console.error("Swap pages failed", err);
          showAlert("ページの入れ替えに失敗しました");
        }
      }
    );
  };

  const handleBulkClearImages = async () => {
    if (isLockedRef.current) return;
    if (selectedSheetIds.size === 0) return;
    const selectedPageCount = selectedSheetIds.size;
    const shouldMoveToTempShelf = selectedPageCount <= 2;

    requestConfirm(
      shouldMoveToTempShelf
        ? `${selectedPageCount}ページ分の画像を外しますか？\n（画像は仮置き場に移動します）`
        : `${selectedPageCount}ページ分の画像を外しますか？\n（画像は未配置リストに戻ります）`,
      async () => {
        // 回収する画像を収集
        const recoveredImages = [];
        const recoveredTempItems = [];
        selectedSheetIds.forEach(id => {
          const sheet = sheets.find(s => s.id === id);
          if (sheet) {
            sheet.panels.forEach(p => {
              const resolvedImage = p.image || (p.imageId ? imageDataById?.[p.imageId] : null);
              if (resolvedImage) {
                const recoveredId = idbHelper.generateId();
                if (shouldMoveToTempShelf) {
                  recoveredTempItems.push({
                    id: recoveredId,
                    ...getPanelTransferableContent(p),
                    image: resolvedImage,
                    imageId: p.imageId || null,
                    originalName: p.code ? `${p.code}.png` : `recovered-${recoveredId}.png`,
                    createdAt: { seconds: Date.now() / 1000 }
                  });
                } else {
                  recoveredImages.push({
                    id: recoveredId,
                    name: p.code ? `${p.code}.png` : `recovered-${recoveredId}.png`,
                    data: resolvedImage,
                    workedBy: undoAccountId ? [undoAccountId] : null,
                    createdAt: { seconds: Date.now() / 1000 }
                  });
                }
              }
            });
          }
        });

        if (USE_LOCAL_STORAGE) {
          const newSheets = sheets.map(s => {
            if (selectedSheetIds.has(s.id)) {
              return {
                ...s,
                panels: s.panels.map(p => ({
                  ...p,
                  image: null,
                  imageId: null,
                  label: null,
                  code: null,
                  text: '',
                  isText: false,
                  ...(shouldMoveToTempShelf ? { freeLabels: [], freeText: null } : {})
                }))
              };
            }
            return s;
          });

          if (shouldMoveToTempShelf) {
            setTempItems(prev => [...recoveredTempItems, ...prev]);
          } else {
            setImages(prev => normalizeStockImages([...prev, ...recoveredImages]));
          }
          setSheets(newSheets);
          // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
          return;
        }

        if (shouldMoveToTempShelf && !tempShelfCollection) {
          showAlert("仮置き場への移動に失敗しました。");
          return;
        }

        const batch = writeBatch(db);
        const savedRecoveredImages = [];

        if (shouldMoveToTempShelf) {
          recoveredTempItems.forEach((item) => {
            const ref = doc(tempShelfCollection);
            batch.set(ref, {
              image: item.image || null,
              imageId: item.imageId || null,
              label: item.label || null,
              code: item.code || null,
              text: item.text || '',
              isText: !!item.isText,
              freeLabels: getPanelFreeLabels(item),
              freeText: null,
              originalName: item.originalName || "退避アイテム",
              ...(useLegacyTempShelf && tempShelfUserId ? { ownerUid: tempShelfUserId } : {}),
              createdAt: serverTimestamp()
            });
          });
        } else {
          // Recovered images to Firestore
          recoveredImages.forEach(img => {
            const ref = doc(imagesCollection);
            savedRecoveredImages.push({ ...img, id: ref.id });
            batch.set(ref, {
              name: img.name,
              data: img.data,
              ...(Array.isArray(img.workedBy) && img.workedBy.length > 0 ? { workedBy: img.workedBy } : {}),
              createdAt: serverTimestamp()
            });
          });
        }

        selectedSheetIds.forEach(id => {
          const sheet = sheets.find(s => s.id === id);
          if (!sheet) return;

          const newPanels = sheet.panels.map(p => ({
            ...p,
            image: null,
            imageId: null,
            label: null,
            code: null,
            text: '',
            isText: false,
            ...(shouldMoveToTempShelf ? { freeLabels: [], freeText: null } : {})
          }));

          const ref = doc(sheetsCollection, id);
          batch.update(ref, { panelsMap: toPanelsMap(newPanels) });
        });

        try {
          await runCloudWrite(() => batch.commit(), { key: 'sheets-bulk-clear' });
          if (!shouldMoveToTempShelf && savedRecoveredImages.length > 0) {
            const nextImages = normalizeStockImages([...images, ...savedRecoveredImages]);
            setImages(nextImages);
            persistCloudImagesCache(nextImages);
          }
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
          if (shouldMoveToTempShelf) {
            showAlert("画像を解除し、仮置き場に移動しました");
          } else {
            showAlert("画像を解除し、未配置リストに戻しました");
          }
        } catch (err) {
          console.error("Bulk clear images failed", err);
          showAlert("一括解除に失敗しました");
        }
      }
    );
  };

  const handleBulkDelete = async () => {
    if (isLockedRef.current) return;
    if (selectedSheetIds.size === 0) return;

    requestConfirm(
      `${selectedSheetIds.size}ページを削除しますか？\n（この操作は取り消せません）`,
      async () => {
        const batch = writeBatch(db);
        selectedSheetIds.forEach(id => {
          const ref = doc(sheetsCollection, id);
          batch.delete(ref);
        });

        try {
          await runCloudWrite(() => batch.commit(), { key: 'sheets-bulk-delete' });
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
        } catch (err) {
          console.error("Bulk delete failed", err);
          showAlert("一括削除に失敗しました");
        }
      }
    );
  };

  // --- CSV Export Logic (for Pages) ---
  const handleExportCSV = () => {
    try {
      const csvContent = buildPageCsvContent({ sheets, genres: GENRES });
      downloadTextFile(csvContent, buildDatedCsvFilename('daiwari_export'));
    } catch (err) {
      console.error("Export failed", err);
      showAlert("CSV出力に失敗しました: " + err.message);
    }
  };

  // --- CSV Import Logic (for Pages) ---
  const handleImportCSV = async (e) => {
    if (isLockedRef.current) return;
    const file = e.target.files[0];
    if (!file) return;

    setIsProcessing(true);
    setProgressValue(0);
    setProgressMax(100);
    setProgressMessage("ファイルを読み込んでいます...");

    try {
      const text = await readFileAutoEncoding(file);
      const rows = text.split(/\r\n|\n|\r/);
      const headers = parseCSVLine(rows[0]);
      if (headers.length < 6) throw new Error('CSVの形式が正しくありません。カラム数が足りません。');

      setProgressMessage("データを解析中...");

      const { sheetUpdates, maxPageIndex } = await parsePageCsvRows(rows, {
        parseLine: parseCSVLine,
        images,
        genres: GENRES,
        onProgress: async (rowIndex, totalRows) => {
          setProgressValue(Math.floor((rowIndex / totalRows) * 100));
          setProgressMessage(`解析中... (${rowIndex}/${totalRows}行)`);
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      });

      // ページを不足分作成 + コマ番号順に再配置
      const finalPageCount = maxPageIndex + 1;
      setProgressMax(finalPageCount);
      const { localSheets, importSummary } = await buildImportedSheets({
        sheets,
        sheetUpdates,
        finalPageCount,
        generateId: () => idbHelper.generateId(),
        onProgress: async (pageIndex) => {
          setProgressValue(pageIndex + 1);
          setProgressMessage(`${pageIndex + 1}ページ目を再構成中...`);
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      });

      let saveFailed = false;
      let saveErrorMessage = '';

      if (!USE_LOCAL_STORAGE) {
        for (let i = 0; i < localSheets.length; i++) {
          const sheet = localSheets[i];
          if (!sheet?.id || !String(sheet.id).startsWith('local_')) continue;

          const newRef = await addDoc(sheetsCollection, {
            genre: sheet.genre || 'none',
            order: sheet.order ?? i,
            panelsMap: toPanelsMap(sheet.panels || buildDefaultPanels()),
            createdAt: serverTimestamp()
          });
          localSheets[i] = { ...sheet, id: newRef.id };
        }
      }

      if (USE_LOCAL_STORAGE) {
        setSheets(localSheets);
        try {
          await idbHelper.setItem('sheets', localSheets);
          setProgressMessage("保存完了");
        } catch (err) {
          console.error("IDB save failed:", err);
          showAlert("自動保存に失敗しました。");
          saveFailed = true;
          saveErrorMessage = err.message;
        }
      } else {
        setSheets(localSheets);
        try {
          let batch = writeBatch(db);
          let opCount = 0;
          for (let i = 0; i < finalPageCount; i++) {
            if (!sheetUpdates[i]) continue;
            const targetSheet = localSheets[i];
            if (!targetSheet?.id) continue;

            if (opCount >= 400) {
              await runCloudWrite(() => batch.commit(), { key: 'csv-import' });
              batch = writeBatch(db);
              opCount = 0;
            }

            batch.update(doc(sheetsCollection, targetSheet.id), {
              genre: targetSheet.genre || 'none',
              order: targetSheet.order ?? i,
              panels: deleteField(),
              panelsMap: toPanelsMap(targetSheet.panels || buildDefaultPanels())
            });
            opCount++;
          }
          if (opCount > 0) {
            await runCloudWrite(() => batch.commit(), { key: 'csv-import' });
          }
        } catch (err) {
          console.error("Cloud save failed:", err);
          saveFailed = true;
          saveErrorMessage = err.message;
        }
      }

      setIsProcessing(false);

      const report = buildPageCsvImportReport(importSummary);

      if (saveFailed) {
        showAlert("保存に失敗したため、取り込み内容は反映されていません。\n" + saveErrorMessage);
      } else {
        showAlert(report, "インポート完了報告", true);
      }

    } catch (err) {
      console.error(err);
      showAlert('エラーが発生しました: ' + err.message);
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const displaySheets = useMemo(() => {
    let result = genreFilter === 'all' ? sheets : sheets.filter(s => s.genre === genreFilter);
    if (viewMode === 'single' && activeSheetId) result = result.filter(s => s.id === activeSheetId);
    return result;
  }, [sheets, genreFilter, viewMode, activeSheetId]);

  const panelArrangeView = useMemo(() => {
    if (!panelArrangeSession) return null;
    const targetSheet = sheets.find((sheet) => sheet.id === panelArrangeSession.sheetId);
    if (!targetSheet?.panels) return null;
    return buildPanelArrangeView(targetSheet.panels, panelArrangeSession);
  }, [panelArrangeSession, sheets]);

  const unresolvedPanelArrangeCount = useMemo(() => (
    getUnresolvedPanelArrangeTokens(panelArrangeView?.session || panelArrangeSession).length
  ), [panelArrangeSession, panelArrangeView]);

  const imageDataById = useMemo(() => {
    const map = {};
    images.forEach((img) => {
      if (img?.id && img?.data) {
        map[img.id] = img.data;
      }
    });
    return map;
  }, [images]);

  const handlePreviewAssignedImage = useCallback((preview) => {
    if (!preview?.src) return;
    const sourceImage = preview.imageId
      ? images.find((image) => image.id === preview.imageId)
      : null;
    const previewName = preview.name || sourceImage?.name || (preview.code ? `${preview.code}.png` : '');
    setAssignedImagePreview({ src: preview.src, name: previewName });
  }, [images]);

  const handleOpenAssignedImage = useCallback((sheetId) => {
    if (panelArrangeSession) {
      showAlert('ホバリングを解除してから別のページへ移動してください。');
      return;
    }
    if (!sheetId || !sheets.some((sheet) => sheet.id === sheetId)) return;
    setGenreFilter('all');
    setActiveSheetId(sheetId);
    setViewMode('single');
    setIsPageSelectionMode(false);
    setSelectedSheetIds(new Set());
    setIsLabelSelectionMode(false);
    setIsMergeMode(false);
    setSelection({ sheetId: null, indices: [] });
  }, [panelArrangeSession, sheets, showAlert]);

  const currentList = useMemo(() => {
    if (viewMode === 'single') return sheets;
    return genreFilter === 'all' ? sheets : sheets.filter(s => s.genre === genreFilter);
  }, [sheets, genreFilter, viewMode]);

  const currentIndex = useMemo(() => {
    return currentList.findIndex(s => s.id === activeSheetId);
  }, [currentList, activeSheetId]);

  const salesLookupVisibleCodes = useMemo(() => {
    if (viewMode !== 'single' || !activeSheetId) return null;
    const activeSheet = sheets.find((sheet) => sheet.id === activeSheetId);
    if (!activeSheet?.panels) return [];

    const codes = new Set();
    activeSheet.panels.forEach((panel) => {
      if (!panel || panel.hidden) return;
      const normalized = normalizeCode(panel.code || '');
      if (normalized) {
        codes.add(normalized);
      }
    });

    return Array.from(codes);
  }, [viewMode, activeSheetId, sheets]);

  const activeSheetLabelCount = useMemo(() => {
    if (viewMode !== 'single' || !activeSheetId) return 0;
    const targetSheet = sheets.find((s) => s.id === activeSheetId);
    if (!targetSheet?.panels) return 0;

    return targetSheet.panels.reduce((count, panel) => {
      const freeLabelsCount = panel?.freeLabels?.length || 0;
      const hasLegacy = !!panel?.freeText && freeLabelsCount === 0;
      return count + freeLabelsCount + (hasLegacy ? 1 : 0);
    }, 0);
  }, [viewMode, activeSheetId, sheets]);

  const handleBulkDeletePageLabels = useCallback(() => {
    if (isLockedRef.current) return;
    if (viewMode !== 'single' || !activeSheetId) return;

    const targetSheet = sheets.find((s) => s.id === activeSheetId);
    if (!targetSheet?.panels) return;

    if (activeSheetLabelCount === 0) {
      showAlert("このページには削除対象のラベルがありません。");
      return;
    }

    requestConfirm(
      `このページの自由ラベルを ${activeSheetLabelCount} 件すべて削除しますか？`,
      async () => {
        const updatedPanels = targetSheet.panels.map((panel) => ({
          ...panel,
          freeLabels: [],
          freeText: null
        }));

        if (USE_LOCAL_STORAGE) {
          setSheets((prev) =>
            prev.map((sheet) =>
              sheet.id === activeSheetId ? { ...sheet, panels: updatedPanels } : sheet
            )
          );
          return;
        }

        try {
          const sheetRef = doc(sheetsCollection, activeSheetId);
          await runCloudTransaction(async (transaction) => {
            const snap = await transaction.get(sheetRef);
            if (!snap.exists()) return;

            const serverPanels = getPanelsFromDocData(snap.data() || {});
            const nextPanels = serverPanels.map((panel) => ({
              ...panel,
              freeLabels: [],
              freeText: null
            }));

            const panelUpdates = buildPanelMapUpdates(serverPanels, nextPanels);
            if (Object.keys(panelUpdates).length > 0) {
              transaction.update(sheetRef, panelUpdates);
            }
          }, { key: `sheet:${activeSheetId}` });
        } catch (err) {
          console.error("Bulk label delete failed", err);
          showAlert("ラベル一括削除に失敗しました。");
        }
      }
    );
  }, [viewMode, activeSheetId, setSheets, sheets, activeSheetLabelCount, sheetsCollection, requestConfirm, showAlert, runCloudTransaction]);

  const handleLogoSecretTap = useCallback(() => {
    logoTapCountRef.current += 1;

    if (logoTapTimeoutRef.current) {
      clearTimeout(logoTapTimeoutRef.current);
    }

    logoTapTimeoutRef.current = setTimeout(() => {
      logoTapCountRef.current = 0;
    }, 1500);

    if (logoTapCountRef.current >= 5) {
      logoTapCountRef.current = 0;
      if (logoTapTimeoutRef.current) {
        clearTimeout(logoTapTimeoutRef.current);
        logoTapTimeoutRef.current = null;
      }
      setIsHiddenImportModalOpen(true);
    }
  }, []);

  const openPageCsvImportFromHiddenMenu = useCallback(() => {
    setIsHiddenImportModalOpen(false);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  }, []);

  const openSalesCsvImportFromHiddenMenu = useCallback(() => {
    setIsHiddenImportModalOpen(false);
    setIsSettingsOpen(true);
  }, []);

  const loadWorkLogDashboard = useCallback(async () => {
    setIsWorkLogLoading(true);
    setWorkLogErrorMessage('');
    try {
      await flushWorkActivityNow();
      if (USE_LOCAL_STORAGE) {
        const stored = JSON.parse(localStorage.getItem(LOCAL_WORK_LOGS_KEY) || '{}');
        setWorkLogRecords(Object.values(stored));
      } else if (workLogsCollection) {
        const snapshot = await getDocs(workLogsCollection);
        setWorkLogRecords(snapshot.docs.map((logDoc) => ({ id: logDoc.id, ...logDoc.data() })));
      }
    } catch (error) {
      console.error('Work log dashboard load failed:', error);
      setWorkLogErrorMessage('作業ログを読み込めませんでした。権限または通信状態を確認してください。');
    } finally {
      setIsWorkLogLoading(false);
    }
  }, [flushWorkActivityNow, workLogsCollection]);

  const openWorkLogDashboardFromHiddenMenu = useCallback(() => {
    setIsHiddenImportModalOpen(false);
    setIsWorkLogDashboardOpen(true);
    void loadWorkLogDashboard();
  }, [loadWorkLogDashboard]);

  const clearQuickHelpHighlight = useCallback(() => {
    const target = quickHelpHighlightRef.current;
    if (!target) return;

    const prevBoxShadow = target.dataset.quickHelpPrevBoxShadow ?? '';
    target.style.boxShadow = prevBoxShadow;
    delete target.dataset.quickHelpPrevBoxShadow;
    quickHelpHighlightRef.current = null;
  }, []);

  const showQuickHelp = useCallback((event, title, description) => {
    if (!isQuickHelpMode) return;
    const target = event.currentTarget;
    if (target && quickHelpHighlightRef.current !== target) {
      clearQuickHelpHighlight();
      target.dataset.quickHelpPrevBoxShadow = target.style.boxShadow || '';
      target.style.boxShadow = '0 0 0 2px rgba(56, 189, 248, 0.75), 0 0 18px rgba(56, 189, 248, 0.45)';
      quickHelpHighlightRef.current = target;
    }

    const rect = event.currentTarget.getBoundingClientRect();
    const popupHalfWidth = 230;
    const x = Math.min(
      Math.max(rect.left + rect.width / 2, popupHalfWidth),
      window.innerWidth - popupHalfWidth
    );

    setQuickHelpPopup({
      title,
      description,
      x,
      y: rect.bottom + 8
    });
  }, [isQuickHelpMode, clearQuickHelpHighlight]);

  const hideQuickHelp = useCallback(() => {
    clearQuickHelpHighlight();
    setQuickHelpPopup(null);
  }, [clearQuickHelpHighlight]);

  useEffect(() => {
    if (!isQuickHelpMode) {
      clearQuickHelpHighlight();
      setQuickHelpPopup(null);
    }
  }, [isQuickHelpMode, clearQuickHelpHighlight]);

  // --- Render ---
  // --- Render ---
  if (!isAuthReady) {
    return (
      <div className="flex h-screen items-center justify-center" style={{ background: 'var(--m3-surface)' }}>
        <div className="flex items-center gap-3 text-sm font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>
          <Loader2 className="h-5 w-5 animate-spin" />
          認証状態を確認中...
        </div>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <AuthGate
        onGoogleSignIn={handleGoogleSignIn}
        isSigningIn={isSigningIn}
        errorMessage={authErrorMessage}
      />
    );
  }

  const handleSelectViewMode = (mode) => {
    if (panelArrangeSession) {
      showAlert('ホバリングを解除してから表示を切り替えてください。');
      return;
    }
    setViewMode(mode);
    setActiveSheetId(null);
    setIsPageSelectionMode(false);
    setIsLabelSelectionMode(false);
  };

  return (
    <div className={`flex flex-col h-screen overflow-hidden transition-all duration-700 ease-in-out`} style={{ background: 'var(--app-bg)', color: 'var(--m3-on-surface)' }}>

      {/* Top Navigation Bar - M3 Expressive Style */}
      {isTopBarsVisible && (
        <AppHeader
          isLocalMode={USE_LOCAL_STORAGE}
          signedInUserName={signedInUserName}
          onLogoTap={handleLogoSecretTap}
          onLogout={handleLogout}
          onUndo={() => { void handleUndoLatest(); }}
          onRedo={() => { void handleRedoLatest(); }}
          viewMode={viewMode}
          onSelectViewMode={handleSelectViewMode}
          isPageSelectionMode={isPageSelectionMode}
          isSalesMode={isSalesMode}
          isSalesLookupOpen={isSalesLookupOpen}
          onSalesModeClick={handleSalesModeButtonClick}
          onSalesModeLongPressStart={startSalesModeLongPress}
          onSalesModeLongPressEnd={endSalesModeLongPress}
          onShowQuickHelp={showQuickHelp}
          onHideQuickHelp={hideQuickHelp}
          selectionToolbar={isPageSelectionMode ? (
            <PageSelectionToolbar
              selectedCount={selectedSheetIds.size}
              totalCount={displaySheets.length}
              isProcessing={isProcessing}
              onSelectAll={handleSelectAllPages}
              onExportPdf={handleExportSelectedPdf}
              onSwap={handleSwapPages}
              onClearImages={handleBulkClearImages}
              onDelete={handleBulkDelete}
            />
          ) : null}
          toolsMenu={(
            <HeaderToolsMenu
              isOpen={isToolsMenuOpen}
              onToggle={() => setIsToolsMenuOpen((prev) => !prev)}
              onClose={() => setIsToolsMenuOpen(false)}
              viewMode={viewMode}
              isLocked={isLocked}
              isPageSelectionMode={isPageSelectionMode}
              isQuickHelpMode={isQuickHelpMode}
              highlightLabels={highlightLabels}
              highlightEmpty={highlightEmpty}
              lockHoldFiredRef={lockHoldFiredRef}
              onStartLockHold={startLockHold}
              onCancelLockHold={cancelLockHold}
              onTogglePageSelectionMode={togglePageSelectionMode}
              onToggleQuickHelpMode={() => setIsQuickHelpMode((prev) => !prev)}
              onToggleHighlightLabels={() => setHighlightLabels(!highlightLabels)}
              onToggleHighlightEmpty={() => setHighlightEmpty(!highlightEmpty)}
              onAddSheet={handleAddSheet}
              onExportCSV={handleExportCSV}
              onShowQuickHelp={showQuickHelp}
              onHideQuickHelp={hideQuickHelp}
            />
          )}
        />
      )}

      <TopBarsToggleButton
        isTopBarsVisible={isTopBarsVisible}
        onToggle={() => {
          setIsTopBarsVisible((current) => !current);
          setIsToolsMenuOpen(false);
          hideQuickHelp();
        }}
      />

      {(viewMode === 'list' || viewMode === 'single') && (
        <>
        <DraggableFloatingPanel
          storageKey="daiwari:floating:sheetControlPanel"
          getDefaultPosition={getSheetControlPanelDefaultPosition}
          className="z-[92] w-40"
        >
        <SheetControlPanel
          viewMode={viewMode}
          isLocked={isLocked}
          isPageSelectionMode={isPageSelectionMode}
          isMergeMode={isMergeMode}
          canMerge={canMerge}
          canSplit={canSplit}
          isLabelSelectionMode={isLabelSelectionMode}
          activeSheetLabelCount={activeSheetLabelCount}
          isPanelArrangeMode={!!panelArrangeSession}
          onToggleMergeMode={toggleMergeMode}
          onMerge={handleMerge}
          onSplit={handleSplit}
          onToggleLabelMode={() => {
            if (panelArrangeSession) {
              showAlert('ホバリング中はラベル追加モードへ切り替えできません。');
              return;
            }
            setIsLabelSelectionMode((current) => !current);
          }}
          onDeleteLabels={handleBulkDeletePageLabels}
          onShowQuickHelp={showQuickHelp}
          onHideQuickHelp={hideQuickHelp}
        />
        </DraggableFloatingPanel>
        <DraggableFloatingPanel
          storageKey="daiwari:floating:tempShelfPanel"
          getDefaultPosition={getTempShelfPanelDefaultPosition}
          className="z-[92] flex w-40 flex-col"
          style={{ maxHeight: '92vh' }}
        >
        <TempShelfPanel
          tempItems={tempItems}
          imageDataById={imageDataById}
          onDeleteFromTemp={handleDeleteFromTemp}
          onApplyDragPayloadToTemp={applyDragPayloadToTempShelf}
          onStartPointerDrag={startPointerDrag}
          onShowQuickHelp={showQuickHelp}
          onHideQuickHelp={hideQuickHelp}
        />
        </DraggableFloatingPanel>
        </>
      )}

      {(viewMode === 'list' || viewMode === 'single') && (
        <ZoomControls zoomScale={zoomScale} setZoomScale={setZoomScale} />
      )}

      <div className="flex flex-1 overflow-hidden relative">
        <Sidebar
          isLocked={isLocked}
          isOpen={sidebarOpen}
          isTopBarsVisible={isTopBarsVisible}
          width={sidebarWidth}
          setWidth={setSidebarWidth}
          toggleOpen={() => setSidebarOpen(!sidebarOpen)}
          images={images}
          sheets={sheets}
          onUpload={handleUploadImage}
          onDeleteImage={handleDeleteImage}
          onBulkDeleteImages={handleBulkDeleteImages}
          onSearch={setSearchQuery}
          searchQuery={searchQuery}
          excludedItems={excludedItems}
          onDeleteFromExcluded={handleDeleteFromExcluded}
          onExportExcludedCSV={handleExportExcludedCSV}
          onBulkDeleteExcluded={handleBulkDeleteExcluded}
          onApplyDragPayloadToExcluded={applyDragPayloadToExcludedList}
          onApplyDragPayloadToStock={applyDragPayloadToStockList}
          onOpenAssignedImage={handleOpenAssignedImage}
          onStartPointerDrag={startPointerDrag}
          imageDataById={imageDataById}
          currentUserUid={undoAccountId}
          onShowQuickHelp={showQuickHelp}
          onHideQuickHelp={hideQuickHelp}
        />

        <div
          className={`flex-1 overflow-auto transition-all relative ${viewMode === 'single' ? 'px-6 pt-1 pb-6' : 'p-8'} ${isSalesMode ? 'text-slate-300' : 'text-slate-800'}`}
          style={{ marginLeft: sidebarOpen ? sidebarWidth : 32 }}
        >
          {/* Background Pattern */}
          <div className="absolute inset-0 z-0 opacity-[0.03] pointer-events-none" style={{ backgroundImage: 'radial-gradient(#64748b 1px, transparent 1px)', backgroundSize: '24px 24px' }}></div>

          {/* Main Content (Sheets) */}
          <div
            className={`relative z-10 flex flex-col ${viewMode === 'single' ? 'gap-4' : 'gap-8'}`}
          >
            {/* Header Controls inside content area */}
            {isTopBarsVisible && (
              <ContentHeaderControls
                viewMode={viewMode}
                genres={GENRES}
                genreFilter={genreFilter}
                onChangeGenreFilter={setGenreFilter}
                currentIndex={currentIndex}
                totalCount={currentList.length}
                activeSheetId={activeSheetId}
                isSalesMode={isSalesMode}
                onNavigate={handleNavigatePage}
              />
            )}

            <div
              className={`relative z-10 ${viewMode === 'overview' ? 'grid grid-cols-2 md:grid-cols-3 gap-8' : 'flex flex-col gap-12 items-center pb-32'}`}
              style={{
                transform: `scale(${zoomScale})`,
                transformOrigin: 'top center',
                minHeight: zoomScale > 1 ? `${zoomScale * 100}%` : 'auto'
              }}
            >
              {displaySheets.map((sheet) => {
                const isPageSelected = selectedSheetIds.has(sheet.id);
                const isArrangeSheet = panelArrangeModeSheetId === sheet.id && !!panelArrangeView;
                const renderedPanels = isArrangeSheet ? panelArrangeView.panels : sheet.panels;
                return (
                  <div
                    key={sheet.id}
                    className={`relative group transition-transform duration-300 ${isPageSelectionMode ? 'cursor-pointer' : ''} ${isPageSelected ? 'scale-[1.02]' : ''}`}
                    onClick={() => {
                      if (isPageSelectionMode) {
                        handleToggleSheetSelection(sheet.id);
                      } else if (viewMode === 'overview') {
                        setActiveSheetId(sheet.id);
                        setIsLabelSelectionMode(false);
                        setViewMode('single');
                      }
                    }}
                  >
                    {isPageSelectionMode && (
                      <div className={`absolute -inset-4 rounded-2xl border-4 z-50 pointer-events-none transition-all duration-200 ${isPageSelected ? 'border-indigo-500 bg-indigo-500/5 shadow-2xl' : 'border-transparent hover:border-slate-300'}`}>
                        <div className={`absolute top-0 right-0 w-8 h-8 rounded-full border-2 bg-white flex items-center justify-center shadow-md transform translate-x-2 -translate-y-2 transition-all ${isPageSelected ? 'border-indigo-500 bg-indigo-500 text-white scale-110' : 'border-slate-300 text-slate-300'}`}>
                          {isPageSelected && <Check size={18} strokeWidth={3} />}
                        </div>
                      </div>
                    )}

                    <div className={`flex flex-col ${viewMode === 'overview' ? 'gap-0' : 'gap-3'}`}>
                      {viewMode !== 'overview' && (
                        <div className="flex items-center justify-between px-2">
                          <span className="font-bold text-slate-500 text-sm flex items-center gap-2">
                            <span className="bg-white border border-slate-200 px-2 py-0.5 rounded text-xs shadow-sm">P.{sheets.findIndex(s => s.id === sheet.id) + 1}</span>
                          </span>

                          <div className="flex items-center gap-2 z-10">
                            <select
                              value={sheet.genre}
                              onChange={(e) => handleChangeGenre(sheet.id, e.target.value)}
                              onClick={(e) => e.stopPropagation()}
                              disabled={isPageSelectionMode}
                              className="text-xs border-none bg-white rounded-lg px-2 py-1 shadow-sm text-slate-600 font-medium focus:ring-2 focus:ring-indigo-500 disabled:opacity-50 hover:bg-slate-50 transition-colors cursor-pointer"
                            >
                              {GENRES.map(g => <option key={g.id} value={g.id}>{g.label}</option>)}
                            </select>
                          </div>
                        </div>
                      )}

                      <div className={`relative ${isPageSelectionMode ? 'pointer-events-none' : ''}`}>
                        {viewMode === 'single' && (
                          <>
                            <button
                              type="button"
                              onClick={(event) => {
                                event.stopPropagation();
                                handleNavigatePage('prev');
                              }}
                              disabled={currentIndex <= 0}
                              className={`absolute left-2 sm:-left-14 top-1/2 z-50 flex h-11 w-11 -translate-y-1/2 items-center justify-center rounded-full border bg-white/95 shadow-lg backdrop-blur transition-all ${currentIndex > 0
                                ? 'border-slate-200 text-slate-600 hover:bg-indigo-50 hover:text-indigo-700 hover:scale-105'
                                : 'border-slate-100 text-slate-300 opacity-40 cursor-not-allowed'
                                }`}
                              title="前のページ"
                              aria-label="前のページ"
                            >
                              <ChevronLeft size={24} />
                            </button>

                            <button
                              type="button"
                              onClick={(event) => {
                                event.stopPropagation();
                                handleNavigatePage('next');
                              }}
                              disabled={currentIndex === -1 || currentIndex >= currentList.length - 1}
                              className={`absolute right-2 sm:-right-14 top-1/2 z-50 flex h-11 w-11 -translate-y-1/2 items-center justify-center rounded-full border bg-white/95 shadow-lg backdrop-blur transition-all ${currentIndex !== -1 && currentIndex < currentList.length - 1
                                ? 'border-slate-200 text-slate-600 hover:bg-indigo-50 hover:text-indigo-700 hover:scale-105'
                                : 'border-slate-100 text-slate-300 opacity-40 cursor-not-allowed'
                                }`}
                              title="次のページ"
                              aria-label="次のページ"
                            >
                              <ChevronRight size={24} />
                            </button>
                          </>
                        )}

                        <Sheet
                          sheet={sheet}
                          index={sheets.findIndex(s => s.id === sheet.id)}
                          pageNumber={sheets.findIndex(s => s.id === sheet.id) + 1}
                          panels={renderedPanels}
                          updatePanel={handlePanelUpdateWithCheck}
                          isOverview={viewMode === 'overview'}
                          zoomScale={zoomScale}
                          selection={selection}
                          onSelectPanel={isMergeMode ? handleSelectPanel : undefined}
                          onDeleteSheet={handleDeleteSheet}
                          highlightEmpty={highlightEmpty}
                          highlightLabels={highlightLabels}
                          onApplyDragPayloadToPanel={applyDragPayloadToPanel}
                          onStartPointerDrag={startPointerDrag}
                          isSalesMode={isSalesMode}
                          salesData={salesData}
                          onHoverSales={handleHoverSales}
                          onLeaveSales={handleLeaveSales}
                          imageDataById={imageDataById}
                          isLabelMode={isLabelSelectionMode}
                          onChangeGenre={(genreId) => handleChangeGenre(sheet.id, genreId)}
                          onPreviewImage={handlePreviewAssignedImage}
                          isArrangeMode={isArrangeSheet}
                          arrangeDraggingTokenId={arrangeDraggingTokenId}
                          arrangeAssignedTokenIdsByPanel={isArrangeSheet ? panelArrangeView.assignedTokenIdsByPanel : {}}
                          arrangePlacedPanelIndices={isArrangeSheet ? panelArrangeView.placedPanelIndices : new Set()}
                          arrangeFloatingTokensByPanel={isArrangeSheet ? panelArrangeView.floatingTokensByPanel : {}}
                          onStartArrangeHold={startPanelArrangeHold}
                          onCancelArrangeHold={clearPanelArrangeHold}
                          onArrangeDragStateChange={handleArrangeDragStateChange}
                        />

                      </div>
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </div>
      </div>

      {/* Global Sales Popup */}
      <SalesPopup
        data={hoveredSalesData}
        position={salesPopupPos}
        onMouseEnter={() => handleHoverSales(hoveredSalesData, null)}
        onMouseLeave={handleLeaveSales}
      />

      <SalesCodeLookupModal
        isOpen={isSalesLookupOpen}
        onClose={() => setIsSalesLookupOpen(false)}
        salesData={salesData}
        visibleCodes={salesLookupVisibleCodes}
      />

      {isQuickHelpMode && quickHelpPopup && <QuickHelpPopup popup={quickHelpPopup} />}

      {panelArrangeModeSheetId && (
        <PanelArrangeBanner
          unresolvedCount={unresolvedPanelArrangeCount}
          isFinalizing={isPanelArrangeFinalizing}
          onFinalize={finalizePanelArrangeMode}
        />
      )}

      {pointerDragPreview && <PointerDragPreview ref={pointerDragOverlayRef} preview={pointerDragPreview} />}

      {undoNotice && <UndoNoticeToast notice={undoNotice} />}

      <input
        type="file"
        accept=".csv"
        ref={fileInputRef}
        onChange={handleImportCSV}
        className="hidden"
      />

      <HiddenImportModal
        isOpen={isHiddenImportModalOpen}
        onClose={() => setIsHiddenImportModalOpen(false)}
        onOpenPageCsvImport={openPageCsvImportFromHiddenMenu}
        onOpenSalesCsvImport={openSalesCsvImportFromHiddenMenu}
        onOpenWorkLogs={openWorkLogDashboardFromHiddenMenu}
      />

      <WorkLogDashboard
        isOpen={isWorkLogDashboardOpen}
        onClose={() => setIsWorkLogDashboardOpen(false)}
        records={workLogRecords}
        isLoading={isWorkLogLoading}
        errorMessage={workLogErrorMessage}
        onRefresh={loadWorkLogDashboard}
        isLocalMode={USE_LOCAL_STORAGE}
      />

      <PdfExportSurface page={pdfExportPage} imageDataById={imageDataById} />

      <ImagePreviewModal preview={assignedImagePreview} onClose={() => setAssignedImagePreview(null)} />

      {/* Settings Modal */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onImportSalesCSV={handleImportSalesCSV}
        salesDataLastUpdated={salesDataLastUpdated}
      />

      <ConfirmModal
        isOpen={confirmDialog.isOpen}
        message={confirmDialog.message}
        onConfirm={confirmDialog.onConfirm}
        onCancel={() => setConfirmDialog({ isOpen: false, message: '', onConfirm: null })}
      />

      <AlertModal
        isOpen={alertDialog.isOpen}
        message={alertDialog.message}
        title={alertDialog.title}
        closeOnBackdrop={alertDialog.closeOnBackdrop}
        onClose={() => setAlertDialog({ ...alertDialog, isOpen: false })}
      />

      <ProcessingModal
        isOpen={isProcessing}
        current={progressValue}
        total={progressMax}
        message={progressMessage}
      />
    </div>
  );
}
