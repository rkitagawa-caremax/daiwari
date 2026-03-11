import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInAnonymously,
  signInWithCustomToken,
  onAuthStateChanged,
  signOut
} from 'firebase/auth';
import {
  getFirestore,
  collection,
  doc,
  setDoc,
  addDoc,
  onSnapshot,
  deleteDoc,
  updateDoc,
  serverTimestamp,
  writeBatch,
  getDoc,
  getDocs,
  runTransaction
} from 'firebase/firestore';
import {
  Plus,
  X,
  Image as ImageIcon,
  Search,
  GripVertical,
  Grid,
  List,
  ChevronLeft,
  ChevronRight,
  Trash2,
  ZoomIn,
  ZoomOut,
  Layout,
  Type,
  Check,
  Merge,
  Split,
  Link as LinkIcon,
  FileSpreadsheet,
  Maximize,
  Upload,
  AlertTriangle,
  Info,
  AlertCircle,
  Globe,
  ClipboardList,
  Copy,
  CheckSquare,
  Square,
  Loader2,
  ArrowLeftRight,
  ArrowLeft,
  ArrowRight,
  Ban,
  Lock,
  Settings,
  BarChart2,
  TrendingUp,
  FileText,
  Database,
  Bell,
  CheckCircle2,
  Tag,
  ChevronDown,
  MoreHorizontal
} from 'lucide-react';

import { idbHelper } from './idbHelper';

// --- Firebase Configuration / Local Storage Mode ---
// Firebase設定 (daiwari-kun)
const firebaseConfig = {
  apiKey: "AIzaSyAMxA79jj3ymqJSCBivjwEfPudnfy8CKAc",
  authDomain: "daiwari-kun.firebaseapp.com",
  projectId: "daiwari-kun",
  storageBucket: "daiwari-kun.firebasestorage.app",
  messagingSenderId: "712325109440",
  appId: "1:712325109440:web:a4dd5d7bcdbb8edf607f25"
};

// 優先順位: 1. グローバル設定があればそれを使用, 2. なければハードコードされた設定を使用
let activeConfig = null;
try {
  activeConfig = typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : firebaseConfig;
} catch (e) {
  activeConfig = firebaseConfig;
}

const resolveStorageMode = () => {
  if (typeof window === 'undefined') return 'cloud';

  try {
    const params = new URLSearchParams(window.location.search);
    const modeParam = (params.get('mode') || '').toLowerCase();
    if (modeParam === 'local' || modeParam === 'cloud') {
      localStorage.setItem('daiwari_storage_mode', modeParam);
      return modeParam;
    }

    const savedMode = (localStorage.getItem('daiwari_storage_mode') || '').toLowerCase();
    if (savedMode === 'local' || savedMode === 'cloud') {
      return savedMode;
    }
  } catch (e) {
    console.warn('Failed to resolve storage mode, defaulting to cloud.', e);
  }

  return 'cloud';
};

let USE_LOCAL_STORAGE = resolveStorageMode() === 'local';

let app, auth, db;
const DEFAULT_APP_ID = typeof __app_id !== 'undefined' ? __app_id : 'default-workspace';

if (!USE_LOCAL_STORAGE) {
  try {
    app = initializeApp(activeConfig);
    auth = getAuth(app);
    db = getFirestore(app);
  } catch (e) {
    console.error('Firebase initialization failed, falling back to localStorage:', e);
    USE_LOCAL_STORAGE = true; // エラー時はローカルモードに強制移行
  }
}

// --- Constants ---
// 視認性を高めるため、カラーコードを濃い色に変更
const GENRES = [
  { id: 'none', label: '未設定', color: '#94a3b8' }, // slate-400
  { id: 'meal', label: '食事関連', color: '#fb923c' }, // orange-400
  { id: 'bath', label: '入浴関連', color: '#facc15' }, // yellow-400
  { id: 'bed', label: 'ベッド', color: '#a3e635' }, // lime-400
  { id: 'clothing', label: '衣類', color: '#4ade80' }, // green-400
  { id: 'walking', label: '歩行関連', color: '#34d399' }, // emerald-400
  { id: 'excretion', label: '排泄関連', color: '#22d3ee' }, // cyan-400
  { id: 'renovation', label: '住宅改修', color: '#60a5fa' }, // blue-400
  { id: 'daily', label: '日常生活', color: '#c084fc' }, // purple-400
  { id: 'health', label: '健康管理', color: '#e879f9' }, // fuchsia-400
  { id: 'medical', label: '医療・施設', color: '#f472b6' }, // pink-400
  { id: 'environment', label: '住環境用品', color: '#fb7185' }, // rose-400
  { id: 'disaster', label: '災害・その他', color: '#f87171' }, // red-400
];

const PASSCODE = "CMC8610";

// --- Free Label Colors ---
const FREE_LABEL_COLORS = [
  { bg: '#34d399', border: '#059669', text: '#ffffff' }, // emerald
  { bg: '#fbbf24', border: '#d97706', text: '#ffffff' }, // amber
  { bg: '#60a5fa', border: '#2563eb', text: '#ffffff' }, // blue
  { bg: '#f472b6', border: '#db2777', text: '#ffffff' }, // pink
  { bg: '#a78bfa', border: '#7c3aed', text: '#ffffff' }, // purple
  { bg: '#fb7185', border: '#e11d48', text: '#ffffff' }, // rose
  { bg: '#22d3ee', border: '#0891b2', text: '#ffffff' }, // cyan
  { bg: '#4ade80', border: '#15803d', text: '#ffffff' }, // green
  { bg: '#f97316', border: '#c2410c', text: '#ffffff' }, // orange
  { bg: '#94a3b8', border: '#475569', text: '#ffffff' }, // slate
];
const FREE_LABEL_HALF_WIDTH_PX = 44;
const FREE_LABEL_HALF_HEIGHT_PX = 20;

const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

const normalizeComparableValue = (value) => (value === undefined ? null : value);

const isPanelDataEqual = (a, b) => {
  if (a === b) return true;
  if (!a || !b) return false;

  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete('fromTempId');
  keys.delete('fromExcludedId');

  for (const key of keys) {
    const av = normalizeComparableValue(a[key]);
    const bv = normalizeComparableValue(b[key]);

    const isObjectLike = (v) => v !== null && typeof v === 'object';
    if (isObjectLike(av) || isObjectLike(bv)) {
      if (JSON.stringify(av) !== JSON.stringify(bv)) return false;
      continue;
    }
    if (av !== bv) return false;
  }
  return true;
};

const hashString = (value = '') => {
  let hash = 0;
  for (let i = 0; i < value.length; i++) {
    hash = ((hash << 5) - hash + value.charCodeAt(i)) | 0;
  }
  return Math.abs(hash).toString(36);
};

const normalizeStockImageEntry = (item, imageDataById = {}) => {
  if (!item) return null;
  const resolvedData = item.data || item.image || (item.imageId ? imageDataById[item.imageId] : null);
  if (!resolvedData) return null;

  const stableId = item.id || item.imageId || `legacy-${hashString(resolvedData)}`;
  return {
    id: stableId,
    name: item.name || item.originalName || item.code || `stock-${stableId}.png`,
    data: resolvedData,
    createdAt: item.createdAt || { seconds: Date.now() / 1000 }
  };
};

const normalizeStockImages = (items = [], imageDataById = {}) => {
  const normalized = [];
  const seenIds = new Set();
  const seenData = new Set();

  items.forEach((item) => {
    const next = normalizeStockImageEntry(item, imageDataById);
    if (!next) return;
    if (seenIds.has(next.id) || seenData.has(next.data)) return;
    seenIds.add(next.id);
    seenData.add(next.data);
    normalized.push(next);
  });

  return normalized;
};

const isSameStockImageList = (a = [], b = []) => {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const left = a[i];
    const right = b[i];
    if ((left?.id || null) !== (right?.id || null)) return false;
    if ((left?.data || null) !== (right?.data || null)) return false;
    if ((left?.name || null) !== (right?.name || null)) return false;
  }
  return true;
};

const getPanelDataPatch = (currentPanel = {}, nextPanel = {}) => {
  const patch = {};
  const keys = new Set([...Object.keys(currentPanel), ...Object.keys(nextPanel)]);

  keys.forEach((key) => {
    const currentValue = normalizeComparableValue(currentPanel[key]);
    const nextValue = normalizeComparableValue(nextPanel[key]);
    const isObjectLike = (v) => v !== null && typeof v === 'object';

    const isChanged = (isObjectLike(currentValue) || isObjectLike(nextValue))
      ? JSON.stringify(currentValue) !== JSON.stringify(nextValue)
      : currentValue !== nextValue;

    if (isChanged) {
      patch[key] = nextPanel[key];
    }
  });

  return patch;
};

const hasPanelTransferableContent = (panel = {}) => {
  return !!(panel.image || panel.imageId || panel.label || panel.isText || panel.code);
};

const clearPanelTransferableContent = (panel = {}) => ({
  ...panel,
  image: null,
  imageId: null,
  label: null,
  code: null,
  text: '',
  isText: false
});

// --- Normalization Utility for Product Codes ---
const normalizeCode = (str) => {
  if (!str) return '';
  return str
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)) // 全角→半角
    .replace(/[-\s]/g, '') // ハイフンやスペースを除去（必要に応じて）
    .trim()
    .toUpperCase(); // 大文字統一
};

// UTF-8 BOM があれば UTF-8、なければ Shift-JIS として自動デコード
const readFileAutoEncoding = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const buf = e.target.result;
    const bytes = new Uint8Array(buf);
    const encoding = (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF)
      ? 'UTF-8' : 'Shift_JIS';
    try {
      const text = new TextDecoder(encoding).decode(buf).replace(/^\uFEFF/, '');
      resolve(text);
    } catch (err) {
      reject(err);
    }
  };
  reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'));
  reader.readAsArrayBuffer(file);
});

// --- Size Definition Helper ---
const getSizeType = (rowSpan, colSpan) => {
  if (rowSpan === 1 && colSpan === 1) return '1/16（1コマ）';
  if (rowSpan === 2 && colSpan === 1) return '1/8 縦（2コマ）';
  if (rowSpan === 1 && colSpan === 2) return '1/8 横（2コマ）';
  if (rowSpan === 3 && colSpan === 1) return '3/16 縦（3コマ）';
  if (rowSpan === 1 && colSpan === 3) return '3/16 横（3コマ）';
  if (rowSpan === 2 && colSpan === 2) return '1/4（4コマ）';
  if (rowSpan === 4 && colSpan === 1) return '1/4 縦（4コマ）';
  if (rowSpan === 1 && colSpan === 4) return '1/4 横（4コマ）';
  if (rowSpan === 3 && colSpan === 2) return '6/16 縦（6コマ）';
  if (rowSpan === 2 && colSpan === 3) return '6/16 横（6コマ）';
  if (rowSpan === 4 && colSpan === 2) return '1/2 縦（8コマ）';
  if (rowSpan === 2 && colSpan === 4) return '1/2 横（8コマ）';
  if (rowSpan === 4 && colSpan === 3) return '12/16 縦（12コマ）';
  if (rowSpan === 4 && colSpan === 4) return '1P（16コマ）';
  return 'custom';
};

// パネルインデックスから4x4グリッドの座標を取得
const getCoords = (index) => ({
  row: Math.floor(index / 4),
  col: index % 4
});

const getSpansFromSizeType = (sizeStr) => {
  if (!sizeStr) return { r: 1, c: 1 };
  // 正規化: 全角英数字を半角に、カッコ・スペース等を削除
  const s = sizeStr
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[（）() \s　]/g, '');

  if (s.includes('1/16')) return { r: 1, c: 1 };
  if (s.includes('1/8縦')) return { r: 2, c: 1 };
  if (s.includes('1/8横')) return { r: 1, c: 2 };
  if (s.includes('3/16縦')) return { r: 3, c: 1 };
  if (s.includes('3/16横')) return { r: 1, c: 3 };
  if (s.includes('1/4縦')) return { r: 4, c: 1 };
  if (s.includes('1/4横')) return { r: 1, c: 4 };
  if (s.includes('1/4')) return { r: 2, c: 2 }; // 正方形の1/4を優先
  if (s.includes('6/16縦')) return { r: 3, c: 2 };
  if (s.includes('6/16横')) return { r: 2, c: 3 };
  if (s.includes('1/2縦')) return { r: 4, c: 2 };
  if (s.includes('1/2横')) return { r: 2, c: 4 };
  if (s.includes('12/16縦')) return { r: 4, c: 3 };
  if (s.includes('12/16横')) return { r: 3, c: 4 }; // 12コマ：3行×4列
  if (s.includes('1P')) return { r: 4, c: 4 };

  return { r: 1, c: 1 };
};

// CSV表記ゆれ耐性版（NFKC正規化 + 縦横判定 + コマ数フォールバック）
const getSpansFromSizeTypeRobust = (sizeStr) => {
  if (!sizeStr) return { r: 1, c: 1 };

  const normalized = String(sizeStr).normalize('NFKC').toLowerCase();
  const s = normalized.replace(/\s+/g, '');
  const hasVertical = /[\u7e26]|vertical/.test(s); // 縦
  const hasHorizontal = /[\u6a2a]|horizontal/.test(s); // 横

  if (s.includes('1/16')) return { r: 1, c: 1 };
  if (s.includes('1/8')) return hasVertical ? { r: 2, c: 1 } : { r: 1, c: 2 };
  if (s.includes('3/16')) return hasVertical ? { r: 3, c: 1 } : { r: 1, c: 3 };
  if (s.includes('1/4')) {
    if (hasVertical) return { r: 4, c: 1 };
    if (hasHorizontal) return { r: 1, c: 4 };
    return { r: 2, c: 2 };
  }
  if (s.includes('6/16')) return hasVertical ? { r: 3, c: 2 } : { r: 2, c: 3 };
  if (s.includes('1/2')) return hasVertical ? { r: 4, c: 2 } : { r: 2, c: 4 };
  if (s.includes('12/16')) return hasVertical ? { r: 4, c: 3 } : { r: 3, c: 4 };
  if (s.includes('1p')) return { r: 4, c: 4 };

  const countMatch = s.match(/(\d+)\s*コマ/);
  const count = countMatch ? parseInt(countMatch[1], 10) : NaN;
  if (count === 1) return { r: 1, c: 1 };
  if (count === 2) return hasVertical ? { r: 2, c: 1 } : { r: 1, c: 2 };
  if (count === 3) return hasVertical ? { r: 3, c: 1 } : { r: 1, c: 3 };
  if (count === 4) {
    if (hasVertical) return { r: 4, c: 1 };
    if (hasHorizontal) return { r: 1, c: 4 };
    return { r: 2, c: 2 };
  }
  if (count === 6) return hasVertical ? { r: 3, c: 2 } : { r: 2, c: 3 };
  if (count === 8) return hasVertical ? { r: 4, c: 2 } : { r: 2, c: 4 };
  if (count === 12) return hasVertical ? { r: 4, c: 3 } : { r: 3, c: 4 };
  if (count === 16) return { r: 4, c: 4 };

  return { r: 1, c: 1 };
};

// --- Utility: Compress Image ---
const compressImage = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = (event) => {
      const img = new Image();
      img.src = event.target.result;
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        const MAX_SIZE = 600;
        if (width > height) {
          if (width > MAX_SIZE) {
            height *= MAX_SIZE / width;
            width = MAX_SIZE;
          }
        } else {
          if (height > MAX_SIZE) {
            width *= MAX_SIZE / height;
            height = MAX_SIZE;
          }
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', 0.7));
      };
      img.onerror = (err) => reject(err);
    };
    reader.onerror = (err) => reject(err);
  });
};

// --- Components ---

// メッセージ用モーダル (Alertの代わり) - Material 3 Style
const AlertModal = React.memo(({ isOpen, message, onClose, title = "通知", closeOnBackdrop = false }) => {
  if (!isOpen) return null;
  return (
    <div
      className="fixed inset-0 z-[90] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in"
      onClick={() => {
        if (closeOnBackdrop) {
          onClose();
        }
      }}
    >
      <div
        className="m3-dialog m3-animate-scale-in"
        style={{ background: 'var(--m3-surface-container-high)' }}
        onClick={(e) => e.stopPropagation()} // ダイアログ内クリックでは閉じないようにする（テキスト選択などのため）
      >
        <div className="flex items-center gap-4 mb-6">
          <div className="p-3 rounded-full" style={{ background: 'var(--m3-primary-container)' }}>
            <Bell size={24} style={{ color: 'var(--m3-on-primary-container)' }} />
          </div>
          <h3 className="text-xl font-medium" style={{ color: 'var(--m3-on-surface)' }}>{title}</h3>
        </div>
        <p className="text-sm mb-8 whitespace-pre-wrap leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>
        <div className="flex justify-end">
          <button
            onClick={onClose}
            className="m3-btn-filled"
          >
            OK
          </button>
        </div>
      </div>
    </div>
  );
});


// 設定モーダル (CSV取り込みなど) - Material 3 Style
const SettingsModal = React.memo(({ isOpen, onClose, onImportSalesCSV, salesDataLastUpdated }) => {
  if (!isOpen) return null;

  const fileInputRef = useRef(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      onImportSalesCSV(file);
    }
  };

  return (
    <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in">
      <div className="m3-dialog w-[520px] overflow-hidden flex flex-col max-h-[90vh] m3-animate-scale-in p-0" style={{ padding: 0 }}>
        <div className="p-5 border-b flex justify-between items-center" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <h3 className="text-lg font-medium flex items-center gap-3" style={{ color: 'var(--m3-on-surface)' }}>
            <div className="p-2 rounded-full" style={{ background: 'var(--m3-secondary-container)' }}>
              <Settings className="w-5 h-5" style={{ color: 'var(--m3-on-secondary-container)' }} />
            </div>
            設定
          </h3>
          <button onClick={onClose} className="m3-icon-btn">
            <X size={20} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-8" style={{ background: 'var(--m3-surface-container-high)' }}>

          {/* 販売数量取り込みセクション */}
          <section>
            <h4 className="text-sm font-medium mb-4 flex items-center gap-3" style={{ color: 'var(--m3-on-surface)' }}>
              <div className="p-1.5 rounded-full" style={{ background: 'var(--m3-tertiary-container)' }}>
                <TrendingUp className="w-4 h-4" style={{ color: 'var(--m3-on-tertiary-container)' }} />
              </div>
              販売数量データの取り込み
            </h4>
            <div className="p-5" style={{ background: 'var(--m3-surface-container-lowest)', borderRadius: 'var(--m3-shape-corner-lg)' }}>
              <p className="text-sm mb-5 leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>
                CSVファイル（商品別売上推移表）を取り込むと、パネル上のコード（介援隊CD）と照合して販売数量を表示できます。<br />
                <span style={{ color: 'var(--m3-error)' }}>※ 取り込みを行うと以前のデータは上書きされます。</span>
              </p>

              <div className="flex items-center gap-4 mb-4">
                <input
                  type="file"
                  accept=".csv"
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="m3-btn-tonal flex items-center gap-2"
                >
                  <FileText size={18} /> CSVを選択して取り込む
                </button>
              </div>

              {salesDataLastUpdated && (
                <div className="flex items-center gap-2 text-xs font-mono px-3 py-2 w-fit" style={{ background: 'var(--m3-surface-container)', borderRadius: 'var(--m3-shape-corner-sm)', color: 'var(--m3-outline)' }}>
                  <Database size={12} />
                  最終更新: {new Date(salesDataLastUpdated).toLocaleString()}
                </div>
              )}
            </div>
          </section>

        </div>

        <div className="p-4 border-t flex justify-end" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <button
            onClick={onClose}
            className="m3-btn-outlined"
          >
            閉じる
          </button>
        </div>
      </div>
    </div>
  );
});

const HiddenImportModal = React.memo(({ isOpen, onClose, onOpenPageCsvImport, onOpenSalesCsvImport }) => {
  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-[85] flex items-center justify-center bg-black/45 backdrop-blur-sm m3-animate-fade-in"
      onClick={onClose}
    >
      <div
        className="m3-dialog w-[420px] overflow-hidden p-0 m3-animate-scale-in"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="p-5 border-b flex justify-between items-center" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
          <h3 className="text-lg font-medium flex items-center gap-2" style={{ color: 'var(--m3-on-surface)' }}>
            <Upload size={18} />
            取込メニュー
          </h3>
          <button onClick={onClose} className="m3-icon-btn">
            <X size={18} />
          </button>
        </div>

        <div className="p-5 space-y-3" style={{ background: 'var(--m3-surface-container-high)' }}>
          <button
            onClick={onOpenPageCsvImport}
            className="w-full m3-btn-tonal flex items-center justify-center gap-2"
          >
            <FileText size={16} />
            台割CSVを取り込む
          </button>

          <button
            onClick={onOpenSalesCsvImport}
            className="w-full m3-btn-tonal flex items-center justify-center gap-2"
          >
            <TrendingUp size={16} />
            販売数量CSVを取り込む
          </button>
        </div>
      </div>
    </div>
  );
});

// 売上詳細ポップアップ
const SalesPopup = React.memo(({ data, position, onMouseEnter, onMouseLeave }) => {
  const popupRef = useRef(null);
  const [offset, setOffset] = useState({ x: 10, y: 10 });

  useEffect(() => {
    if (popupRef.current && position) {
      const rect = popupRef.current.getBoundingClientRect();
      const viewportWidth = window.innerWidth;
      const viewportHeight = window.innerHeight;

      let nextX = 10;
      let nextY = 10;

      if (position.x + rect.width + 20 > viewportWidth) {
        nextX = -rect.width - 10;
      }
      if (position.y + rect.height + 20 > viewportHeight) {
        nextY = viewportHeight - (position.y + rect.height + 20);
      }
      setOffset({ x: nextX, y: nextY });
    }
  }, [position, data]);

  if (!data || !position) return null;

  const totalCount = data.reduce((sum, item) => sum + (parseInt(item.count) || 0), 0);

  return (
    <div
      ref={popupRef}
      className="fixed z-[100] bg-white/95 backdrop-blur-md rounded-2xl shadow-2xl border border-slate-200 p-4 w-72 animate-in fade-in zoom-in-95 duration-200 pointer-events-auto"
      style={{
        top: position.y + offset.y,
        left: position.x + offset.x,
        boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1), 0 0 0 1px rgb(0 0 0 / 0.05)'
      }}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
    >
      <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-3">
        <div className="flex flex-col">
          <span className="text-[10px] uppercase tracking-wider font-bold text-slate-400">Sales Record</span>
          <span className="text-xs font-bold text-slate-600">販売実績詳細</span>
        </div>
        <div className="flex flex-col items-end">
          <span className="text-2xl font-black text-indigo-600 font-mono leading-none">{totalCount.toLocaleString()}</span>
          <span className="text-[10px] text-slate-400 font-bold uppercase">Total Units</span>
        </div>
      </div>
      <ul className="space-y-2 max-h-64 overflow-y-auto pr-1 custom-scrollbar">
        {data.map((item, idx) => (
          <li key={idx} className="text-[11px] leading-tight flex justify-between gap-3 py-2 border-b border-slate-50 last:border-0 hover:bg-slate-50/50 transition-colors rounded-lg px-1">
            <div className="flex flex-col gap-0.5 min-w-0">
              <span className="font-bold text-slate-800 truncate">{item.name}</span>
              <span className="text-slate-500 truncate text-[9px] opacity-80">{item.spec}</span>
            </div>
            <span className="font-mono font-bold text-indigo-500 bg-indigo-50 px-1.5 py-0.5 rounded ml-auto self-center">{item.count}</span>
          </li>
        ))}
      </ul>
    </div>
  );
});

// 処理中モーダル - Material 3 Style
const ProcessingModal = React.memo(({ isOpen, current, total, message }) => {
  if (!isOpen) return null;
  const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/50 backdrop-blur-sm cursor-wait m3-animate-fade-in">
      <div className="m3-dialog w-96 flex flex-col items-center m3-animate-scale-in">
        <div className="relative mb-6">
          <div className="absolute inset-0 rounded-full blur-xl opacity-30 animate-pulse" style={{ background: 'var(--m3-primary)' }}></div>
          <Loader2 className="w-14 h-14 animate-spin relative z-10" style={{ color: 'var(--m3-primary)' }} />
        </div>
        <h3 className="text-xl font-medium mb-2" style={{ color: 'var(--m3-on-surface)' }}>処理中...</h3>
        <p className="text-sm mb-6 text-center whitespace-pre-wrap" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>

        <div className="w-full h-1 rounded-full mb-3 overflow-hidden" style={{ background: 'var(--m3-surface-container-highest)' }}>
          <div
            className="h-1 rounded-full transition-all duration-300 ease-out"
            style={{ width: `${percentage}%`, background: 'var(--m3-primary)' }}
          ></div>
        </div>
        <div className="flex justify-between w-full text-xs font-mono" style={{ color: 'var(--m3-outline)' }}>
          <span>{current} / {total}</span>
          <span>{percentage}%</span>
        </div>
        <p className="text-xs mt-6 font-medium flex items-center gap-2" style={{ color: 'var(--m3-error)' }}>
          <AlertTriangle size={14} /> 画面を閉じないでください
        </p>
      </div>
    </div>
  );
});

const AuthGate = ({ onAuthenticated, defaultAppId }) => {
  const [input, setInput] = useState("");
  const [workspaceId, setWorkspaceId] = useState(defaultAppId);
  const [error, setError] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (input === PASSCODE) {
      onAuthenticated(workspaceId);
    } else {
      setError(true);
      setInput("");
    }
  };

  return (
    <div className="flex items-center justify-center min-h-screen relative overflow-hidden" style={{ background: 'var(--m3-surface)' }}>
      {/* Expressive background shapes */}
      <div className="absolute top-[-15%] left-[-10%] w-[50%] h-[50%] rounded-full blur-3xl opacity-40" style={{ background: 'var(--m3-primary-container)' }} />
      <div className="absolute bottom-[-15%] right-[-10%] w-[45%] h-[45%] rounded-full blur-3xl opacity-40" style={{ background: 'var(--m3-tertiary-container)' }} />
      <div className="absolute top-[30%] right-[20%] w-[20%] h-[20%] rounded-full blur-2xl opacity-30" style={{ background: 'var(--m3-secondary-container)' }} />

      <div className="m3-card-elevated p-10 w-[420px] relative z-10 m3-animate-scale-in" style={{ borderRadius: 'var(--m3-shape-corner-xl)' }}>
        <div className="flex justify-center mb-10">
          <div className="p-1 bg-white shadow-xl" style={{ borderRadius: 'var(--m3-shape-corner-xl)' }}>
            <img src="/logo.jpg" alt="台割君" className="w-40 h-40 object-contain" style={{ borderRadius: 'calc(var(--m3-shape-corner-xl) - 4px)' }} />
          </div>
        </div>
        <p className="text-base mb-10 text-center font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>アクセスコードを入力してください</p>

        <form onSubmit={handleSubmit} className="space-y-6">
          <div className="space-y-2">
            <label className="text-xs font-medium ml-1" style={{ color: 'var(--m3-on-surface-variant)' }}>Workspace ID</label>
            <div className="relative group">
              <Globe className="absolute left-4 top-4 w-5 h-5 transition-colors" style={{ color: 'var(--m3-outline)' }} />
              <input
                type="text"
                value={workspaceId}
                onChange={(e) => setWorkspaceId(e.target.value)}
                className="w-full p-4 pl-12 font-medium transition-all m3-input-outlined"
                style={{ borderRadius: 'var(--m3-shape-corner-md)' }}
                placeholder="Project Name / ID"
              />
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-xs font-medium ml-1" style={{ color: 'var(--m3-on-surface-variant)' }}>Passcode</label>
            <div className="relative group">
              <Lock className="absolute left-4 top-4 w-5 h-5 transition-colors" style={{ color: 'var(--m3-outline)' }} />
              <input
                type="password"
                value={input}
                onChange={(e) => setInput(e.target.value)}
                className="w-full p-4 pl-12 font-mono tracking-widest placeholder:tracking-normal transition-all m3-input-outlined"
                style={{ borderRadius: 'var(--m3-shape-corner-md)' }}
                placeholder="CMC8610"
              />
            </div>
          </div>

          {error && (
            <div className="flex items-center justify-center gap-3 text-sm p-4 m3-animate-fade-in" style={{ background: 'var(--m3-error-container)', color: 'var(--m3-on-error-container)', borderRadius: 'var(--m3-shape-corner-md)' }}>
              <AlertCircle size={18} />
              <span className="font-medium">パスコードが違います</span>
            </div>
          )}
          <button
            type="submit"
            className="w-full p-4 font-medium text-base transition-all duration-200 hover:shadow-lg active:scale-[0.98]"
            style={{
              background: 'var(--m3-primary)',
              color: 'var(--m3-on-primary)',
              borderRadius: 'var(--m3-shape-corner-full)',
              boxShadow: 'var(--m3-elevation-2)'
            }}
          >
            ロック解除
          </button>
        </form>
      </div>
    </div>
  );
};

// 確認用モーダル - Material 3 Style
const ConfirmModal = React.memo(({ isOpen, message, onConfirm, onCancel }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/50 backdrop-blur-sm m3-animate-fade-in">
      <div className="m3-dialog m3-animate-scale-in" style={{ background: 'var(--m3-surface-container-high)' }}>
        <div className="flex items-center gap-4 mb-6">
          <div className="p-3 rounded-full" style={{ background: 'var(--m3-error-container)' }}>
            <AlertTriangle size={24} style={{ color: 'var(--m3-on-error-container)' }} />
          </div>
          <h3 className="text-xl font-medium" style={{ color: 'var(--m3-on-surface)' }}>確認</h3>
        </div>
        <p className="text-sm mb-8 whitespace-pre-wrap leading-relaxed" style={{ color: 'var(--m3-on-surface-variant)' }}>{message}</p>
        <div className="flex justify-end gap-3">
          <button
            onClick={onCancel}
            className="m3-btn-text"
          >
            キャンセル
          </button>
          <button
            onClick={onConfirm}
            className="px-6 py-2.5 font-medium transition-all active:scale-[0.98]"
            style={{
              background: 'var(--m3-error)',
              color: 'var(--m3-on-error)',
              borderRadius: 'var(--m3-shape-corner-full)'
            }}
          >
            実行する
          </button>
        </div>
      </div>
    </div>
  );
});

// Panel Component
const Panel = React.memo(({ index, data, globalNumber, onUpdate, onUpdateByIndex, panels, isOverview, isSelected, onSelect, highlightEmpty, highlightLabels, sheetId, onMove, isSalesMode, salesData, onHoverSales, onLeaveSales, imageDataById, isLabelMode }) => {

  const [isHovered, setIsHovered] = useState(false);
  const textareaRef = useRef(null);
  const [localText, setLocalText] = useState(data.text || '');
  const [labelDrafts, setLabelDrafts] = useState({});
  const isFocusedRef = useRef(false);
  const editingLabelIdRef = useRef(null);
  const panelRef = useRef(null);


  // 売上データマッチング
  const matchedSales = useMemo(() => {
    if (!isSalesMode || !data.code || !salesData) return null;
    const normalizedTarget = normalizeCode(data.code);
    return salesData[normalizedTarget] || null;
  }, [isSalesMode, data.code, salesData]);

  // Firestoreデータと同期
  useEffect(() => {
    if (!isFocusedRef.current) {
      setLocalText(data.text || '');
    }
  }, [data.text]);

  useEffect(() => {
    const labels = data.freeLabels || (data.freeText ? [{ id: 'legacy', text: data.freeText, x: 50, y: 50, colorIndex: 0 }] : []);
    setLabelDrafts((prev) => {
      const next = {};
      labels.forEach((label) => {
        if (editingLabelIdRef.current === label.id && prev[label.id] !== undefined) {
          next[label.id] = prev[label.id];
        } else {
          next[label.id] = label.text || '';
        }
      });
      return next;
    });
  }, [data.freeLabels, data.freeText]);




  const handleMouseEnter = (e) => {
    setIsHovered(true);
    if (isSalesMode && matchedSales && onHoverSales) {
      // パネルの位置を取得して親に渡す
      const rect = e.currentTarget.getBoundingClientRect();
      onHoverSales(matchedSales, { x: rect.right, y: rect.top });
    }
  };

  const handleMouseLeave = () => {
    setIsHovered(false);
    if (isSalesMode && onLeaveSales) {
      onLeaveSales();
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    if (isOverview) return;

    // Check if it's a move operation (Panel to Panel)
    const moveSourceType = e.dataTransfer.getData("moveSourceType");
    if (moveSourceType === "panel") {
      const sourceSheetId = e.dataTransfer.getData("sourceSheetId");
      const sourceIndex = parseInt(e.dataTransfer.getData("sourceIndex"), 10);
      const movedText = e.dataTransfer.getData("textData"); // 移動元のテキストデータ

      if (sourceSheetId && !isNaN(sourceIndex) && onMove) {
        onMove(sourceSheetId, sourceIndex, sheetId, index, movedText);
      }
      return;
    }

    // Normal Drop
    const type = e.dataTransfer.getData("type");
    let src = e.dataTransfer.getData("src");
    if (src === "null" || src === "undefined") src = "";

    let droppedImageId = e.dataTransfer.getData("imageId");
    if (droppedImageId === "null" || droppedImageId === "undefined" || !droppedImageId) droppedImageId = null;

    let label = e.dataTransfer.getData("label");
    if (label === "null" || label === "undefined" || !label) label = null;

    let droppedCode = e.dataTransfer.getData("code");
    if (droppedCode === "null" || droppedCode === "undefined" || !droppedCode) droppedCode = null;

    const fileName = e.dataTransfer.getData("name");
    const isText = e.dataTransfer.getData("isText");
    const hasTextPayload = e.dataTransfer.getData("hasTextPayload") === "1";
    const transferredText = hasTextPayload
      ? e.dataTransfer.getData("textPayload")
      : e.dataTransfer.getData("text");

    let fromTempId = e.dataTransfer.getData("fromTempId");
    if (fromTempId === "null" || fromTempId === "undefined" || !fromTempId) fromTempId = null;

    let fromExcludedId = e.dataTransfer.getData("fromExcludedId");
    if (fromExcludedId === "null" || fromExcludedId === "undefined" || !fromExcludedId) fromExcludedId = null;

    if (src || droppedImageId || label || droppedCode || isText === "true" || fromTempId || fromExcludedId) {
      let code = droppedCode || null;

      if (!code && fileName && !isText) {
        const match = fileName.match(/[A-Za-z]\d{4}/);
        if (match) {
          code = match[0].toUpperCase();
        }
      }

      const initialText = isText === "true"
        ? ((hasTextPayload || transferredText !== "") ? transferredText : (data.text || "テキストを入力"))
        : "";

      onUpdate({
        ...data,
        image: src || null,
        imageId: droppedImageId || null,
        label: label || null,
        code: code,
        isText: isText === "true",
        text: initialText,
        fromTempId: fromTempId || null,
        fromExcludedId: fromExcludedId || null
      });
      setLocalText(initialText);
    }
  };

  const handleDragStart = (e) => {
    if (!resolvedImage && !data.label && !data.isText && !data.code) {
      e.preventDefault();
      return;
    }

    const currentText = textareaRef.current ? textareaRef.current.value : localText;
    e.dataTransfer.setData("moveSourceType", "panel");
    e.dataTransfer.setData("sourceSheetId", sheetId);
    e.dataTransfer.setData("sourceIndex", index);
    // 編集中のテキストも転送データに含める
    e.dataTransfer.setData("textData", currentText || "");
    e.dataTransfer.effectAllowed = "move";
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
  };

  const resolvedImage = data.image || (data.imageId ? imageDataById?.[data.imageId] : null);
  const isEmpty = !resolvedImage && !data.label && !data.code && !data.isText;
  const freeLabelsCount = data.freeLabels?.length || 0;
  const hasFreeLabel = freeLabelsCount > 0 || (!!data.freeText && freeLabelsCount === 0);
  const shouldHighlightLabel = isOverview && highlightLabels && hasFreeLabel;
  const shouldHighlightEmpty = highlightEmpty && (!resolvedImage && (isEmpty || !!data.code));

  const handleFocusTextarea = () => {
    if (textareaRef.current) {
      textareaRef.current.focus();
    }
  };

  const handleTextChange = (e) => {
    setLocalText(e.target.value);
  };

  const handleTextBlur = () => {
    isFocusedRef.current = false;
    if (localText !== data.text) {
      onUpdate({ ...data, text: localText });
    }
  };

  const getLabelStyle = (label) => {
    switch (label) {
      case '新規商品':
      case '新規商品未確定': return { bg: 'var(--dummy-gray)', text: 'var(--dummy-text-color)' };
      case 'テキスト': return { bg: 'var(--dummy-volt)', text: 'var(--dummy-text-color)' };
      case 'タイトル': return { bg: 'var(--dummy-red)', text: 'var(--dummy-text-color)' };
      case '埋草': return { bg: 'var(--dummy-green)', text: 'var(--dummy-text-color)' };
      default: return { bg: 'var(--m3-surface-container)', text: 'var(--m3-on-surface)' };
    }
  };

  const handlePanelClick = (e) => {
    if (isOverview) return;

    if (isLabelMode) {
      e.stopPropagation();
      if (!panelRef.current) return;

      const currentRect = panelRef.current.getBoundingClientRect();
      let targetIndex = index;
      let targetPanelEl = panelRef.current;

      const overflowDirections = [];
      if ((e.clientX - currentRect.left) < FREE_LABEL_HALF_WIDTH_PX) overflowDirections.push('left');
      if ((currentRect.right - e.clientX) < FREE_LABEL_HALF_WIDTH_PX) overflowDirections.push('right');
      if ((e.clientY - currentRect.top) < FREE_LABEL_HALF_HEIGHT_PX) overflowDirections.push('up');
      if ((currentRect.bottom - e.clientY) < FREE_LABEL_HALF_HEIGHT_PX) overflowDirections.push('down');

      if (overflowDirections.length > 0 && typeof document !== 'undefined') {
        const idPrefix = `panel-${sheetId}-`;
        const viewportMaxX = Math.max(0, window.innerWidth - 1);
        const viewportMaxY = Math.max(0, window.innerHeight - 1);

        for (const direction of overflowDirections) {
          let probeX = e.clientX;
          let probeY = e.clientY;

          if (direction === 'left') probeX = currentRect.left - 1;
          if (direction === 'right') probeX = currentRect.right + 1;
          if (direction === 'up') probeY = currentRect.top - 1;
          if (direction === 'down') probeY = currentRect.bottom + 1;

          const candidate = document.elementFromPoint(
            clamp(probeX, 0, viewportMaxX),
            clamp(probeY, 0, viewportMaxY)
          );
          const panelEl = candidate?.closest?.(`[id^="${idPrefix}"]`);
          if (!panelEl || panelEl === panelRef.current) continue;

          const rawIndex = (panelEl.id || '').slice(idPrefix.length);
          const parsedIndex = Number.parseInt(rawIndex, 10);
          if (Number.isNaN(parsedIndex)) continue;
          if (panels?.[parsedIndex]?.hidden) continue;

          targetIndex = parsedIndex;
          targetPanelEl = panelEl;
          break;
        }
      }

      const targetRect = targetPanelEl.getBoundingClientRect();
      const rawXPercent = ((e.clientX - targetRect.left) / targetRect.width) * 100;
      const rawYPercent = ((e.clientY - targetRect.top) / targetRect.height) * 100;

      const minXPercent = (FREE_LABEL_HALF_WIDTH_PX / targetRect.width) * 100;
      const maxXPercent = 100 - minXPercent;
      const minYPercent = (FREE_LABEL_HALF_HEIGHT_PX / targetRect.height) * 100;
      const maxYPercent = 100 - minYPercent;

      const xPercent = minXPercent <= maxXPercent ? clamp(rawXPercent, minXPercent, maxXPercent) : 50;
      const yPercent = minYPercent <= maxYPercent ? clamp(rawYPercent, minYPercent, maxYPercent) : 50;

      const targetPanelData = targetIndex === index ? data : (panels?.[targetIndex] || {});
      const prevLabels = targetPanelData.freeLabels || [];
      const migratedLabels = targetPanelData.freeText && prevLabels.length === 0
        ? [{ id: 'legacy', text: targetPanelData.freeText, x: 50, y: 50, colorIndex: 0 }]
        : prevLabels;

      const colorIndex = migratedLabels.length % FREE_LABEL_COLORS.length;

      const newLabel = {
        id: Date.now().toString() + Math.random().toString(36).substring(2, 7),
        text: 'ラベル',
        x: xPercent,
        y: yPercent,
        colorIndex
      };

      const nextPanelData = {
        ...targetPanelData,
        freeText: null,
        freeLabels: [...migratedLabels, newLabel]
      };

      if (targetIndex === index || !onUpdateByIndex) {
        onUpdate(nextPanelData);
      } else {
        onUpdateByIndex(targetIndex, nextPanelData);
      }
      return;
    }

    if (onSelect) onSelect();
  };

  const labelStyle = data.label ? getLabelStyle(data.label) : {};


  // 売上合計
  const salesTotal = matchedSales ? matchedSales.reduce((acc, item) => acc + (parseInt(item.count) || 0), 0) : 0;

  return (
    <div
      ref={panelRef}
      id={`panel-${sheetId}-${index}`}
      className={`relative border-t border-l flex flex-col items-center justify-center overflow-hidden transition-all duration-300
        ${(shouldHighlightLabel || shouldHighlightEmpty) ? 'ring-inset ring-2' : 'hover:shadow-md hover:z-10'} 
        ${isSelected ? 'ring-4 z-20 shadow-xl' : ''}
        ${!isEmpty && !isOverview && !data.isText ? 'cursor-grab active:cursor-grabbing' : ''}
        ${isSalesMode ? 'hover:ring-4 hover:z-40' : ''}
      `}
      draggable={!isEmpty && !isOverview && !isLabelMode && !data.isText}
      onDragStart={handleDragStart}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      onClick={handlePanelClick}
      style={{
        ... (shouldHighlightEmpty ? {} : {}),
        '--tw-ring-color': shouldHighlightLabel
          ? '#22c55e'
          : shouldHighlightEmpty
            ? 'var(--m3-error)'
            : 'transparent',
        cursor: isLabelMode
          ? 'crosshair'
          : (isOverview ? 'pointer' : (!isEmpty && !data.isText ? 'cursor-grab' : 'default')),
        gridColumn: `span ${data.colSpan || 1}`,
        gridRow: `span ${data.rowSpan || 1}`,
        borderColor: shouldHighlightLabel ? '#16a34a' : 'var(--m3-outline-variant)',
        backgroundColor: shouldHighlightLabel
          ? 'rgba(34, 197, 94, 0.16)'
          : shouldHighlightEmpty
          ? 'var(--m3-error-container)'
          : isSalesMode && matchedSales
            ? 'var(--m3-secondary-container)'
            : 'var(--m3-surface)',
        boxShadow: shouldHighlightLabel
          ? '0 0 0 1.5px rgba(22, 163, 74, 0.65), inset 0 0 0 1px rgba(34, 197, 94, 0.55), 0 0 20px rgba(34, 197, 94, 0.35)'
          : undefined,
      }}
    >

      {/* Image Layer */}
      {resolvedImage && (
        <img src={resolvedImage} alt="content" className="w-full h-full object-contain absolute inset-0 z-0 pointer-events-none" loading="lazy" decoding="async" />
      )}

      {/* Free Labels Layer (Overlays everything) */}
      {(() => {
        const labels = data.freeLabels || (data.freeText ? [{ id: 'legacy', text: data.freeText, x: 50, y: 50, colorIndex: 0 }] : []);
        if (labels.length === 0) return null;

        return labels.map((label, i) => {
          const color = FREE_LABEL_COLORS[label.colorIndex % FREE_LABEL_COLORS.length];
          const draftText = labelDrafts[label.id] ?? label.text ?? '';
          return (
            <div
              key={label.id}
              className="absolute z-[25] transform -translate-x-1/2 -translate-y-1/2"
              style={{ left: `${label.x}%`, top: `${label.y}%`, minWidth: '80px', maxWidth: '90%' }}
              onClick={(e) => e.stopPropagation()}
              onMouseDown={(e) => e.stopPropagation()} // ドラッグ誤動作防止
            >
              <div className="relative group flex items-center justify-center">
                <textarea
                  className="w-full bg-white/95 backdrop-blur-sm rounded-lg p-1.5 text-xs font-bold text-center resize-none focus:outline-none shadow-lg min-h-[2.5em] overflow-hidden transition-shadow focus:ring-2"
                  style={{
                    border: `2px solid ${color.border}`,
                    color: '#334155', // slate-700
                    '--tw-ring-color': color.border
                  }}
                  value={draftText}
                  onChange={(e) => {
                    const nextText = e.target.value;
                    setLabelDrafts((prev) => ({ ...prev, [label.id]: nextText }));
                  }}
                  onFocus={() => { editingLabelIdRef.current = label.id; }}
                  onBlur={(e) => {
                    const committedText = (e.target.value || '').toString();
                    editingLabelIdRef.current = null;
                    if (committedText === (label.text || '')) return;
                    const newLabels = [...labels];
                    newLabels[i] = { ...label, text: committedText };
                    onUpdate({ ...data, freeLabels: newLabels, freeText: null });
                  }}
                  placeholder="入力"
                  rows={Math.max(1, draftText.split('\n').length)}
                />
                {!isOverview && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      const newLabels = labels.filter((_, idx) => idx !== i);
                      onUpdate({ ...data, freeLabels: newLabels, freeText: null });
                    }}
                    className="absolute -top-2 -right-2 text-white rounded-full p-1 shadow-md transition-colors opacity-0 group-hover:opacity-100 group-focus-within:opacity-100 z-10"
                    style={{ backgroundColor: color.border }}
                    title="ラベルを削除"
                  >
                    <X size={12} strokeWidth={3} />
                  </button>
                )}
              </div>
            </div>
          );
        });
      })()}


      {/* Sales Overlay Mode */}
      {isSalesMode && matchedSales && (
        <div className="absolute inset-0 z-30 bg-black/60 flex flex-col p-2 text-white pointer-events-none">
          <div className="flex justify-between items-start mb-1">
            <span className="text-[10px] bg-emerald-500 text-white px-1 py-0.5 rounded font-bold shadow-sm">
              実績
            </span>
            <span className="text-xl font-bold font-mono tracking-tighter text-emerald-300">
              {salesTotal.toLocaleString()}
            </span>
          </div>
          {/* 簡易リスト表示 (最大3件) */}
          <div className="flex-1 overflow-hidden space-y-1">
            {matchedSales.slice(0, 3).map((item, i) => (
              <div key={i} className="flex justify-between items-baseline text-[9px] border-b border-white/20 pb-0.5">
                <span className="truncate w-2/3 opacity-90">{item.name} {item.spec}</span>
                <span className="font-mono font-bold opacity-100">{item.count}</span>
              </div>
            ))}
            {matchedSales.length > 3 && (
              <div className="text-[8px] text-center opacity-70 italic mt-1">
                他 {matchedSales.length - 3} 件...
              </div>
            )}
          </div>
        </div>
      )}

      {/* Drop Placeholder */}
      {!resolvedImage && (!data.code || (highlightEmpty && (!resolvedImage && (isEmpty || !!data.code)))) && !data.label && !isOverview && (
        <div className={`flex flex-col items-center justify-center transition-opacity duration-300 ${(highlightEmpty && (!resolvedImage && (isEmpty || !!data.code))) ? 'opacity-100' : 'opacity-0 hover:opacity-100'}`}>
          <span className={`text-[10px] select-none font-bold`} style={{ color: (highlightEmpty && isEmpty) ? 'var(--m3-on-error-container)' : 'var(--m3-outline)' }}>
            {highlightEmpty && isEmpty ? '空き' : 'Drop Here'}
          </span>
        </div>
      )}

      {/* Label/Overlay for Dummy Images (Non-text) */}
      {data.label && !data.isText && (
        <div
          className="absolute inset-0 flex items-center justify-center pointer-events-none opacity-90 z-10"
          style={{ background: labelStyle.bg, color: labelStyle.text }}
        >
          <span className="font-bold text-sm">{data.label}</span>
        </div>
      )}

      {/* Text Area for Text Dummy */}
      {data.isText && (
        isOverview ? (
          <div
            className="absolute inset-0 p-2 z-20 flex items-center justify-center text-center overflow-hidden pointer-events-none"
            style={{ background: labelStyle.bg || 'rgba(255,255,255,0.5)' }}
          >
            <p className="text-[8px] leading-tight break-words whitespace-pre-wrap font-bold font-sans" style={{ color: labelStyle.text || 'var(--m3-on-surface)' }}>{data.text}</p>
          </div>
        ) : (
          <div
            data-text-editor="true"
            className="absolute inset-0 z-20 flex items-center justify-center p-4 cursor-text transition-colors hover:brightness-95 focus-within:brightness-100"
            style={{ background: labelStyle.bg || 'rgba(255,255,255,0.4)' }}
            onClick={(e) => { e.stopPropagation(); handleFocusTextarea(); }}
            onMouseDown={(e) => e.stopPropagation()}
          >
            <button
              className="absolute top-1 left-1 rounded-md p-1.5 cursor-grab active:cursor-grabbing border shadow-sm z-30"
              style={{ background: 'var(--m3-surface)', borderColor: 'var(--m3-outline-variant)', color: 'var(--m3-on-surface-variant)' }}
              title="ドラッグして移動"
              draggable
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => e.stopPropagation()}
              onDragStart={handleDragStart}
            >
              <GripVertical size={12} />
            </button>
            <textarea
              ref={textareaRef}
              className={`w-full bg-transparent resize-none focus:outline-none text-sm font-bold text-center overflow-hidden font-sans placeholder:text-slate-400/70 ${labelStyle.text || 'text-slate-800'}`}
              value={localText}
              onChange={handleTextChange}
              onFocus={() => isFocusedRef.current = true}
              onBlur={handleTextBlur}
              placeholder="テキストを入力"
              rows={Math.max(2, (localText || '').split('\n').length)}
              style={{ maxHeight: '100%' }}
              draggable={false}
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => e.stopPropagation()}
            />
          </div>
        )
      )}

      {/* Numbering & Code Badge */}
      {globalNumber && !isSalesMode && (
        <div className="absolute top-0 left-0 z-10 flex shadow-sm pointer-events-none opacity-90">
          <div className="text-[10px] px-1.5 py-0.5 font-bold rounded-br-sm shadow-sm" style={{ background: 'var(--m3-surface-variant)', color: 'var(--m3-on-surface-variant)' }}>
            {globalNumber}
          </div>
          {data.code && (
            <div className="text-[10px] px-1.5 py-0.5 font-bold border-r border-b shadow-sm font-mono" style={{ background: 'var(--m3-surface)', color: 'var(--m3-on-surface)', borderColor: 'var(--m3-outline-variant)' }}>
              {data.code}
            </div>
          )}
        </div>
      )}

      {/* Controls (Only in List Mode) */}
      {!isOverview && (
        <>
          {(resolvedImage || data.label || data.isText || data.code) && isHovered && !isSalesMode && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                onUpdate({ ...data, image: null, imageId: null, label: null, code: null, isText: false, text: '' });
              }}
              className="absolute top-1.5 right-1.5 rounded-full p-1.5 shadow-lg border z-30 transition-all duration-300 transform hover:scale-110 active:scale-90"
              style={{ background: 'var(--m3-surface)', color: 'var(--m3-error)', borderColor: 'var(--m3-error)' }}
            >
              <X size={14} strokeWidth={3} />
            </button>
          )}
        </>
      )}
    </div>
  );
});

const Sheet = React.memo(({ sheet, index, panels, updatePanel, isOverview, zoomScale, selection, onSelectPanel, onDeleteSheet, highlightEmpty, highlightLabels, onMovePanel, isSalesMode, salesData, onHoverSales, onLeaveSales, imageDataById, isLabelMode }) => {

  const genre = GENRES.find(g => g.id === sheet.genre) || GENRES[0];

  let visibleCounter = 0;
  const displayNumbers = {};
  for (let i = 0; i < 16; i++) {
    const p = panels[i];
    if (!p?.hidden) {
      const isNonCounted = p.label === '埋草' || p.label === 'タイトル';

      if (!isNonCounted) {
        visibleCounter++;
        displayNumbers[i] = visibleCounter;
      } else {
        displayNumbers[i] = null;
      }
    }
  }

  return (
    <div
      className={`border transition-all duration-300 overflow-hidden ${isOverview ? 'hover:shadow-lg hover:scale-[1.02] cursor-pointer' : 'shadow-xl'}`}
      style={{
        width: isOverview ? '100%' : `${210 * zoomScale}mm`,
        height: isOverview ? 'auto' : `${297 * zoomScale}mm`,
        aspectRatio: '210/297',
        margin: '0 auto',
        display: 'flex',
        flexDirection: 'column',
        borderRadius: 'var(--m3-shape-corner-lg)',
        background: 'var(--m3-surface)',
        borderColor: 'var(--m3-outline-variant)'
      }}
    >
      {/* Sheet Header */}
      <div
        className="h-3 w-full flex-shrink-0"
        style={{ backgroundColor: genre.color }}
        title={genre.label}
      />

      {/* Grid Content */}
      <div className="flex-1 grid grid-cols-4 grid-rows-4 w-full h-full border-b border-r" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}>
        {Array.from({ length: 16 }).map((_, i) => {
          const panelData = panels[i] || {};
          if (panelData.hidden) return null;

          const panelNumber = displayNumbers[i];
          const isSelected = !isOverview && selection?.sheetId === sheet.id && selection.indices.includes(i);

          return (
            <Panel
              key={i}
              index={i}
              data={panelData}
              globalNumber={panelNumber}
              onUpdate={(newData) => updatePanel(sheet.id, i, newData)}
              onUpdateByIndex={(targetIndex, newData) => updatePanel(sheet.id, targetIndex, newData)}
              panels={panels}
              isOverview={isOverview}
              isSelected={isSelected}
              onSelect={() => onSelectPanel && onSelectPanel(sheet.id, i)}
              highlightEmpty={highlightEmpty}
              highlightLabels={highlightLabels}
              sheetId={sheet.id}
              onMove={onMovePanel}
              isSalesMode={isSalesMode}
              salesData={salesData}
              onHoverSales={onHoverSales}
              onLeaveSales={onLeaveSales}
              imageDataById={imageDataById}
              isLabelMode={isLabelMode}
            />

          );
        })}
      </div>

      {/* Overview Footer */}
      {isOverview && (
        <div className="text-center text-[10px] py-1.5 border-t flex justify-between items-center px-3 font-medium" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)', color: 'var(--m3-on-surface-variant)' }}>
          <span className="truncate">Page {index + 1} - {genre.label}</span>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onDeleteSheet(sheet.id);
            }}
            className="p-1.5 rounded-full transition-colors hover:bg-black/5"
            style={{ color: 'var(--m3-on-surface-variant)' }}
            title="削除"
          >
            <Trash2 size={14} />
          </button>
        </div>
      )}
    </div>
  );
});

// ... Sidebar Component ...
const Sidebar = React.memo(({ isOpen, width, setWidth, toggleOpen, images, onUpload, onDeleteImage, onBulkDeleteImages, onSearch, searchQuery, sheets, tempItems, onMoveToTemp, onDeleteFromTemp, excludedItems, onMoveToExcluded, onDeleteFromExcluded, onExportExcludedCSV, onBulkDeleteExcluded, onMoveToStock, imageDataById, onShowQuickHelp, onHideQuickHelp }) => {
  const [activeTab, setActiveTab] = useState('stock');
  const [resizing, setResizing] = useState(false);
  const [statusGenreFilter, setStatusGenreFilter] = useState('all');
  const [isImageSelectionMode, setIsImageSelectionMode] = useState(false);
  const [selectedImageIds, setSelectedImageIds] = useState(new Set());

  const sidebarTabHelp = {
    stock: { title: '画像', description: '画像ライブラリを表示します。画像の検索・アップロード・配置ができます。' },
    dummy: { title: 'ダミー', description: 'ダミー素材をドラッグしてコマに配置します。テキストダミーもここから使えます。' },
    status: { title: '空き', description: 'ページごとの空きコマ状況を確認できます。ジャンル絞り込みにも対応しています。' },
    excluded: { title: '除外', description: '掲載除外リストです。除外した素材の確認・復帰・削除ができます。' }
  };

  const getImageSelectionKey = useCallback((img) => {
    if (img?.id) return `id:${img.id}`;
    const data = img?.data || '';
    const head = data.slice(0, 32);
    const tail = data.slice(-32);
    return `legacy:${img?.name || ''}:${data.length}:${head}:${tail}`;
  }, []);

  useEffect(() => {
    setSelectedImageIds(new Set());
    setIsImageSelectionMode(false);
  }, [activeTab]);

  const toggleImageSelection = (selectionKey) => {
    setSelectedImageIds(prev => {
      const newSet = new Set(prev);
      if (newSet.has(selectionKey)) newSet.delete(selectionKey);
      else newSet.add(selectionKey);
      return newSet;
    });
  };

  const handleBulkDelete = () => {
    const selectedTargets = filteredImages
      .filter((img) => selectedImageIds.has(getImageSelectionKey(img)))
      .map((img) => ({ id: img.id || null, data: img.data || null }));
    onBulkDeleteImages(selectedTargets);
    setIsImageSelectionMode(false);
    setSelectedImageIds(new Set());
  };

  useEffect(() => {
    const handleMouseMove = (e) => {
      if (resizing) {
        const newWidth = Math.max(200, Math.min(600, e.clientX));
        setWidth(newWidth);
      }
    };
    const handleMouseUp = () => setResizing(false);

    if (resizing) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizing, setWidth]);

  const handleDropToTemp = (e) => {
    e.preventDefault();
    const moveSourceType = e.dataTransfer.getData("moveSourceType");
    if (moveSourceType === "panel") {
      const sourceSheetId = e.dataTransfer.getData("sourceSheetId");
      const sourceIndex = parseInt(e.dataTransfer.getData("sourceIndex"), 10);
      const movedText = e.dataTransfer.getData("textData");

      if (sourceSheetId && !isNaN(sourceIndex)) {
        onMoveToTemp(sourceSheetId, sourceIndex, movedText);
      }
    }
  };

  const handleDropToStock = (e) => {
    e.preventDefault();
    const moveSourceType = e.dataTransfer.getData("moveSourceType");
    if (moveSourceType === "panel") {
      const sourceSheetId = e.dataTransfer.getData("sourceSheetId");
      const sourceIndex = parseInt(e.dataTransfer.getData("sourceIndex"), 10);
      if (sourceSheetId && !isNaN(sourceIndex) && onMoveToStock) {
        onMoveToStock(sourceSheetId, sourceIndex);
      }
    }
  };

  const handleDropToExcluded = (e) => {
    e.preventDefault();
    const moveSourceType = e.dataTransfer.getData("moveSourceType");
    if (moveSourceType === "panel") {
      const sourceSheetId = e.dataTransfer.getData("sourceSheetId");
      const sourceIndex = parseInt(e.dataTransfer.getData("sourceIndex"), 10);
      const movedText = e.dataTransfer.getData("textData");

      if (sourceSheetId && !isNaN(sourceIndex)) {
        onMoveToExcluded(sourceSheetId, sourceIndex, movedText);
      }
    }
  };

  // 画像検索フィルタリング (配置済み画像を除外)
  const filteredImages = useMemo(() => {
    // 全シートで使用されている画像のSetを作成
    const usedImageSet = new Set();
    sheets.forEach(sheet => {
      sheet.panels.forEach(p => {
        if (p.image) usedImageSet.add(p.image);
        if (p.imageId) usedImageSet.add(p.imageId);
      });
    });

    let result = images.filter(img => !usedImageSet.has(img.data) && !usedImageSet.has(img.id));

    if (!searchQuery) return result;
    const lowerQuery = searchQuery.toLowerCase();
    return result.filter(img => (img.name || '').toLowerCase().includes(lowerQuery));
  }, [images, searchQuery, sheets]);

  if (!isOpen) {
    return (
      <div className="fixed left-0 top-16 bottom-0 w-16 m3-surface border-r flex flex-col items-center py-6 z-20 transition-all duration-300" style={{ borderColor: 'var(--m3-outline-variant)' }}>
        <button
          onClick={toggleOpen}
          className="m3-icon-btn-tonal p-3"
          style={{ background: 'var(--m3-primary-container)', color: 'var(--m3-on-primary-container)' }}
        >
          <ChevronRight size={24} />
        </button>
      </div>
    );
  }

  return (
    <div
      className="fixed left-0 top-16 bottom-0 m3-surface border-r flex flex-col z-20 shadow-xl transition-all duration-300 ease-in-out"
      style={{ width, borderColor: 'var(--m3-outline-variant)' }}
    >
      <div className="flex items-center justify-between px-4 py-4 border-b flex-shrink-0" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
        <div className="flex p-1 rounded-full flex-1 mr-4 overflow-x-auto" style={{ background: 'var(--m3-surface-container-highest)' }}>
          {[
            { id: 'stock', icon: ImageIcon, label: '画像' },
            { id: 'dummy', icon: Type, label: 'ダミー' },
            { id: 'status', icon: Info, label: '空き' },
            { id: 'excluded', icon: Ban, label: '除外' }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              onMouseEnter={(e) => {
                const help = sidebarTabHelp[tab.id];
                if (help) onShowQuickHelp?.(e, help.title, help.description);
              }}
              onMouseLeave={() => onHideQuickHelp?.()}
              className={`flex-1 flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-full transition-all duration-200 whitespace-nowrap ${activeTab === tab.id ? 'bg-white shadow-sm' : 'hover:bg-black/5 opacity-70'}`}
              style={activeTab === tab.id ? { color: 'var(--m3-primary)' } : { color: 'var(--m3-on-surface-variant)' }}
            >
              <tab.icon size={16} />
              <span className="text-xs font-medium">{tab.label}</span>
            </button>
          ))}
        </div>
        <button onClick={toggleOpen} className="m3-icon-btn flex-shrink-0">
          <ChevronLeft size={24} />
        </button>
      </div>

      {activeTab === 'stock' && (
        <div className="p-4 border-b flex-shrink-0" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}>
          <div className="relative group">
            <Search className="absolute left-4 top-3.5 w-5 h-5 transition-colors" style={{ color: 'var(--m3-outline)' }} />
            <input
              type="text"
              placeholder="画像検索..."
              value={searchQuery}
              onChange={(e) => onSearch(e.target.value)}
              onMouseEnter={(e) => onShowQuickHelp?.(e, '画像検索', '画像名でライブラリを絞り込みます。キーワード入力で候補を素早く探せます。')}
              onMouseLeave={() => onHideQuickHelp?.()}
              className="w-full p-3 pl-12 rounded-full transition-all"
              style={{ background: 'var(--m3-surface-container-high)', color: 'var(--m3-on-surface)' }}
            />
            {searchQuery && (
              <button
                onClick={() => onSearch('')}
                className="absolute right-3 top-3 p-1 rounded-full hover:bg-black/10 transition-colors"
                style={{ color: 'var(--m3-on-surface-variant)' }}
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>
      )}

      <div
        className="flex-1 overflow-y-auto p-4 relative"
        style={{ background: 'var(--m3-surface-container-low)' }}
        onDragOver={(e) => e.preventDefault()}
        onDrop={activeTab === 'stock' ? handleDropToStock : undefined}
      >

        {activeTab === 'stock' && (
          <div className="space-y-6">
            <div className="space-y-2">
              <label
                className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl cursor-pointer transition-all group"
                style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}
                onMouseEnter={(e) => onShowQuickHelp?.(e, 'クリックしてアップロード', '画像ファイルを選択してライブラリに追加します。複数選択アップロードにも対応します。')}
                onMouseLeave={() => onHideQuickHelp?.()}
              >
                <div className="flex flex-col items-center pt-2 pb-3">
                  <div className="p-3 rounded-full mb-3 transition-colors" style={{ background: 'var(--m3-primary-container)' }}>
                    <ImageIcon className="w-6 h-6" style={{ color: 'var(--m3-on-primary-container)' }} />
                  </div>
                  <p className="text-sm font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>クリックしてアップロード</p>
                </div>
                <input type="file" className="hidden" accept="image/*" multiple onChange={onUpload} />
              </label>
            </div>

            <div className="flex justify-between items-center px-1">
              <span className="text-sm font-bold flex items-center gap-2" style={{ color: 'var(--m3-on-surface)' }}>
                {searchQuery ? '検索結果' : 'ライブラリ'}
                <span className="px-2 py-0.5 rounded-full text-xs" style={{ background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' }}>{filteredImages.length}</span>
              </span>
              <button
                onClick={() => {
                  setIsImageSelectionMode(!isImageSelectionMode);
                  setSelectedImageIds(new Set());
                }}
                onMouseEnter={(e) => onShowQuickHelp?.(e, 'ライブラリのチェックボックス', '複数画像を選択するモードです。ONで一括選択・一括削除が使えます。')}
                onMouseLeave={() => onHideQuickHelp?.()}
                className={`p-2 rounded-full transition-all ${isImageSelectionMode ? 'bg-primary-container text-on-primary-container ring-1' : ''}`}
                style={isImageSelectionMode ? { background: 'var(--m3-primary-container)', color: 'var(--m3-on-primary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
                title="画像を選択して削除"
              >
                <CheckSquare size={20} />
              </button>
            </div>

            {isImageSelectionMode && (
              <div className="flex items-center justify-between p-4 rounded-xl shadow-sm m3-animate-fade-in" style={{ background: 'var(--m3-surface-container-high)' }}>
                <span className="text-sm font-bold" style={{ color: 'var(--m3-primary)' }}>{selectedImageIds.size}枚選択中</span>
                <div className="flex gap-2">
                  <button
                    onClick={() => setSelectedImageIds(new Set(filteredImages.map((i) => getImageSelectionKey(i))))}
                    className="text-xs font-medium px-3 py-1.5 rounded-full hover:bg-black/5"
                    style={{ color: 'var(--m3-primary)' }}
                  >
                    全選択
                  </button>
                  <button
                    onClick={handleBulkDelete}
                    disabled={selectedImageIds.size === 0}
                    className={`text-xs px-4 py-1.5 rounded-full font-medium transition-all ${selectedImageIds.size > 0 ? 'shadow-sm' : 'opacity-50 cursor-not-allowed'}`}
                    style={{ background: 'var(--m3-error)', color: 'var(--m3-on-error)' }}
                  >
                    削除
                  </button>
                </div>
              </div>
            )}

            <div className="grid grid-cols-2 gap-3" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))' }}>
              {filteredImages.map((img) => {
                const selectionKey = getImageSelectionKey(img);
                return (
                <div
                  key={selectionKey}
                  className={`group relative border rounded-xl p-2 transition-all duration-200
                    ${isImageSelectionMode
                      ? 'cursor-pointer'
                      : 'cursor-grab active:cursor-grabbing hover:shadow-md hover:-translate-y-0.5'} 
                    ${selectedImageIds.has(selectionKey)
                      ? 'ring-2'
                      : ''}`}
                  style={{
                    background: 'var(--m3-surface)',
                    borderColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'var(--m3-outline-variant)',
                    backgroundColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary-container)' : 'var(--m3-surface)'
                  }}
                  draggable={!isImageSelectionMode}
                  onClick={() => {
                    if (isImageSelectionMode) {
                      toggleImageSelection(selectionKey);
                    }
                  }}
                  onDragStart={(e) => {
                    if (isImageSelectionMode) {
                      e.preventDefault();
                      return;
                    }
                    e.dataTransfer.setData("src", img.data);
                    if (img.id) e.dataTransfer.setData("imageId", img.id);
                    e.dataTransfer.setData("type", "image");
                    e.dataTransfer.setData("name", img.name);
                  }}
                >
                  <div className="aspect-square w-full rounded-lg overflow-hidden mb-2 bg-white flex items-center justify-center">
                    <img src={img.data} alt="stock" className="max-w-full max-h-full object-contain" loading="lazy" decoding="async" />
                  </div>
                  <div className="px-1">
                    <div className="text-[10px] truncate font-medium" style={{ color: 'var(--m3-on-surface)' }}>{img.name}</div>
                  </div>

                  {isImageSelectionMode ? (
                    <div className="absolute top-2 left-2 w-6 h-6 rounded-full flex items-center justify-center transition-all"
                      style={{ background: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'rgba(255,255,255,0.8)', border: selectedImageIds.has(selectionKey) ? 'none' : '1px solid var(--m3-outline)' }}>
                      {selectedImageIds.has(selectionKey) && <Check size={14} style={{ color: 'var(--m3-on-primary)' }} />}
                    </div>
                  ) : (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteImage(img.id, img.data);
                      }}
                      className="absolute top-2 right-2 rounded-full p-2 opacity-0 group-hover:opacity-100 shadow-sm transition-all hover:scale-110"
                      style={{ background: 'var(--m3-surface)', color: 'var(--m3-error)', border: '1px solid var(--m3-outline-variant)' }}
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
                );
              })}

              {filteredImages.length === 0 && (
                <div className="col-span-full flex flex-col items-center justify-center py-16 text-center" style={{ color: 'var(--m3-outline)' }}>
                  <Search size={32} className="mb-3 opacity-20" />
                  <p className="text-sm">{searchQuery ? '見つかりませんでした' : '画像がありません'}</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Dummy Tab */}
        {activeTab === 'dummy' && (
          <div className="space-y-4">
            <p className="text-xs font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>ドラッグして配置できます</p>

            {[
              { color: 'var(--dummy-gray)', text: 'var(--dummy-text-color)', label: '新規商品未確定' },
              { color: 'var(--dummy-red)', text: 'var(--dummy-text-color)', label: 'タイトル' },
              { color: 'var(--dummy-green)', text: 'var(--dummy-text-color)', label: '埋草' },
              { color: 'var(--dummy-volt)', text: 'var(--dummy-text-color)', label: 'テキスト', isText: true },
            ].map((dummy) => (
              <div
                key={dummy.label}
                className="h-24 border-2 border-dashed rounded-xl flex items-center justify-center cursor-grab active:cursor-grabbing hover:shadow-md hover:-translate-y-0.5 transition-all"
                style={{ background: dummy.color, borderColor: 'var(--m3-outline-variant)', color: dummy.text }}
                draggable
                onDragStart={(e) => {
                  const canvas = document.createElement('canvas');
                  canvas.width = 100; canvas.height = 100;
                  const ctx = canvas.getContext('2d');
                  ctx.fillStyle = '#eef2ff';
                  ctx.fillRect(0, 0, 100, 100);
                  e.dataTransfer.setData("src", canvas.toDataURL());
                  e.dataTransfer.setData("label", dummy.label);
                  if (dummy.isText) e.dataTransfer.setData("isText", "true");
                }}
              >
                <span className="font-bold text-sm">【{dummy.label}】</span>
              </div>
            ))}
          </div>
        )}


        {/* Status Tab */}
        {activeTab === 'status' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="font-bold text-slate-600 flex items-center gap-2 text-xs">
                <Info size={16} /> 空き状況
              </h3>
              <select
                value={statusGenreFilter}
                onChange={(e) => setStatusGenreFilter(e.target.value)}
                className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 text-slate-600"
              >
                <option value="all">全ジャンル</option>
                {GENRES.map(g => (
                  <option key={g.id} value={g.id}>{g.label}</option>
                ))}
              </select>
            </div>

            <div className="space-y-2">
              {sheets
                .map((sheet, originalIndex) => ({ sheet, originalIndex }))
                .filter(({ sheet }) => statusGenreFilter === 'all' || sheet.genre === statusGenreFilter)
                .map(({ sheet, originalIndex }) => {
                  const pureEmptyCount = sheet.panels.filter(p => !p.hidden && !p.image && !p.imageId && !p.label).length;
                  const dummyCount = sheet.panels.filter(p => !p.hidden && p.label && p.label !== 'タイトル' && p.label !== '埋草' && p.label !== 'テキスト').length;
                  const totalEmpty = pureEmptyCount + dummyCount;

                  const dummyDetails = sheet.panels.reduce((acc, p) => {
                    if (!p.hidden && p.label && p.label !== 'タイトル' && p.label !== '埋草') {
                      acc[p.label] = (acc[p.label] || 0) + 1;
                    }
                    return acc;
                  }, {});

                  const genre = GENRES.find(g => g.id === sheet.genre) || GENRES[0];

                  return (
                    <div key={sheet.id} className="flex flex-col bg-white border border-slate-100 rounded-lg p-3 shadow-sm hover:shadow-md transition-shadow">
                      <div className="flex justify-between items-center mb-2">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-slate-700 text-xs">P{originalIndex + 1}</span>
                          <span
                            className="text-[10px] px-2 py-0.5 rounded-full font-medium truncate max-w-[100px]"
                            style={{
                              backgroundColor: genre.color,
                              color: '#1e293b' // Darker text for better contrast
                            }}
                          >
                            {genre.label}
                          </span>
                        </div>
                        <span className={`font-mono text-xs font-bold ${totalEmpty > 0 ? 'text-rose-500' : 'text-emerald-500'}`}>
                          {totalEmpty > 0 ? `空き: ${totalEmpty}` : '完了'}
                        </span>
                      </div>

                      {Object.keys(dummyDetails).length > 0 && (
                        <div className="flex flex-wrap gap-1.5 pl-4 border-l-2 border-slate-200 ml-1">
                          {Object.entries(dummyDetails).map(([label, count]) => {
                            const dummyAttr = [
                              { color: 'bg-[var(--dummy-gray)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: '新規商品未確定' },
                              { color: 'bg-[var(--dummy-orange)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'その他' },
                              { color: 'bg-[var(--dummy-red)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'タイトル' },
                              { color: 'bg-[var(--dummy-green)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: '埋草' },
                              { color: 'bg-[var(--dummy-volt)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'テキスト' },
                            ].find(d => d.label === label) || { color: 'bg-slate-50', text: 'text-slate-500', border: 'border-slate-200' };

                            return (
                              <span key={label} className={`text-[9.5px] px-2 py-0.5 ${dummyAttr.color} ${dummyAttr.text} rounded-full border ${dummyAttr.border} font-bold shadow-sm`}>
                                {label}: {count}
                              </span>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              {sheets.length === 0 && <p className="text-xs text-slate-400 text-center py-4">ページがありません</p>}
            </div>
          </div>
        )}
        {activeTab === 'excluded' && (
          <div
            className="flex-1 flex flex-col h-full bg-rose-50/30 rounded-xl border border-rose-100 overflow-hidden"
            onDragOver={(e) => e.preventDefault()}
            onDrop={handleDropToExcluded}
          >
            <div className="p-3 border-b border-rose-100 flex items-center justify-between bg-white/50">
              <div className="flex items-center gap-2 text-sm font-bold text-rose-600">
                <Ban size={16} />
                <span>掲載除外リスト</span>
              </div>
              <div className="flex items-center gap-1">
                <button
                  onClick={onExportExcludedCSV}
                  disabled={excludedItems.length === 0}
                  className={`flex items-center gap-1.5 px-2 py-1 text-xs rounded-md border transition-colors font-medium ${excludedItems.length > 0 ? 'bg-white text-rose-600 border-rose-200 hover:bg-rose-50' : 'bg-slate-50 text-slate-300 border-transparent cursor-not-allowed'}`}
                  title="除外リストをCSVで出力"
                >
                  <FileSpreadsheet size={12} />
                </button>
                <button
                  onClick={onBulkDeleteExcluded}
                  disabled={excludedItems.length === 0}
                  className={`flex items-center gap-1.5 px-2 py-1 text-xs rounded-md border transition-colors font-medium ${excludedItems.length > 0 ? 'bg-white text-rose-600 border-rose-200 hover:bg-rose-500 hover:text-white' : 'bg-slate-50 text-slate-300 border-transparent cursor-not-allowed'}`}
                  title="除外リストを全て空にする"
                >
                  <Trash2 size={12} />
                </button>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-2">
              {excludedItems.length === 0 ? (
                <div className="h-full flex flex-col items-center justify-center text-rose-300 py-8 border-2 border-dashed border-rose-200/50 rounded-lg m-1">
                  <Ban size={24} className="mb-2 opacity-50" />
                  <p className="text-[10px]">アイテムがありません</p>
                  <p className="text-[9px] opacity-70">ここへドロップして除外</p>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  {excludedItems.map((item) => {
                    const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                    return (
                      <div
                        key={item.id}
                        className="group relative border border-rose-100 rounded-lg p-2 bg-white hover:shadow-md cursor-grab active:cursor-grabbing flex flex-col items-center transition-all"
                        draggable
                        onDragStart={(e) => {
                          const payloadText = typeof item.text === 'string' ? item.text : '';
                          e.dataTransfer.setData("src", resolvedImg || "");
                          e.dataTransfer.setData("type", "image");
                          e.dataTransfer.setData("name", item.originalName || "excluded");
                          e.dataTransfer.setData("label", item.label || "");
                          e.dataTransfer.setData("code", item.code || "");
                          e.dataTransfer.setData("isText", item.isText ? "true" : "false");
                          e.dataTransfer.setData("hasTextPayload", "1");
                          e.dataTransfer.setData("textPayload", payloadText);
                          e.dataTransfer.setData("text", payloadText);
                          e.dataTransfer.setData("fromExcludedId", item.id);
                          if (item.imageId) e.dataTransfer.setData("imageId", item.imageId);
                        }}
                      >
                        <div className="w-full aspect-square bg-slate-50 rounded mb-2 overflow-hidden">
                          {resolvedImg ? (
                            <img src={resolvedImg} alt="excluded" className="w-full h-full object-contain" />
                          ) : (
                            <div className="w-full h-full flex flex-col items-center justify-center text-slate-300">
                              <Ban size={16} />
                            </div>
                          )}
                        </div>

                        <div className="w-full flex justify-between items-center text-[9px]">
                          <span className="font-bold text-rose-500 truncate max-w-[60px] font-mono">{item.code || 'No Code'}</span>
                          {item.label && (
                            <span className="bg-slate-100 text-slate-500 px-1 rounded truncate max-w-[50px]">{item.label}</span>
                          )}
                        </div>

                        <button
                          onClick={() => onDeleteFromExcluded(item.id)}
                          className="absolute -top-1.5 -right-1.5 bg-white rounded-full p-1 shadow-sm text-rose-400 hover:text-rose-600 hover:bg-rose-50 border border-rose-100 opacity-0 group-hover:opacity-100 transition-all"
                          title="完全に削除"
                        >
                          <X size={12} strokeWidth={3} />
                        </button>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Temp Shelf (Fixed at bottom) */}
      <div
        className="h-64 border-t border-slate-200 bg-slate-50 flex flex-col flex-shrink-0"
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDropToTemp}
      >
        <div
          className="px-4 py-2 border-b border-slate-200 flex items-center justify-between bg-white"
          onMouseEnter={(e) => onShowQuickHelp?.(e, '仮置き場', 'コマを一時退避する場所です。ドラッグで戻せるので並べ替えや調整に使えます。')}
          onMouseLeave={() => onHideQuickHelp?.()}
        >
          <div className="flex items-center gap-2 text-xs font-bold text-slate-600">
            <div className="p-1 bg-indigo-100 text-indigo-600 rounded">
              <ClipboardList size={14} />
            </div>
            <span>仮置き場</span>
          </div>
          <span className="text-[10px] text-slate-400 font-medium">ドラッグして一時保存</span>
        </div>

        <div className="flex-1 overflow-y-auto p-3">
          {tempItems.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-400 border-2 border-dashed border-slate-200 rounded-xl">
              <p className="text-xs font-medium">ここにドロップ</p>
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-3">
              {tempItems.map((item) => {
                const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                return (
                  <div
                    key={item.id}
                    className="group relative border border-slate-200 rounded-lg p-2 bg-white hover:shadow-md cursor-grab active:cursor-grabbing flex items-center justify-center min-h-[80px] transition-all"
                    draggable
                    onDragStart={(e) => {
                      const payloadText = typeof item.text === 'string' ? item.text : '';
                      e.dataTransfer.setData("src", resolvedImg || "");
                      e.dataTransfer.setData("type", "image");
                      e.dataTransfer.setData("name", item.originalName || "temp");
                      e.dataTransfer.setData("label", item.label || "");
                      e.dataTransfer.setData("code", item.code || "");
                      e.dataTransfer.setData("isText", item.isText ? "true" : "false");
                      e.dataTransfer.setData("hasTextPayload", "1");
                      e.dataTransfer.setData("textPayload", payloadText);
                      e.dataTransfer.setData("text", payloadText);
                      e.dataTransfer.setData("fromTempId", item.id);
                      if (item.imageId) e.dataTransfer.setData("imageId", item.imageId);
                    }}
                  >
                    {resolvedImg ? (
                      <img src={resolvedImg} alt="temp" className="w-full h-16 object-contain rounded" />
                    ) : (
                      <div className="w-full h-16 flex flex-col items-center justify-center bg-slate-50 rounded text-slate-400">
                        <span className="text-[10px] font-mono">{item.code || 'No Image'}</span>
                      </div>
                    )}

                    {item.label && (
                      <div className="absolute top-1 left-1 bg-slate-800/80 backdrop-blur-sm text-white text-[9px] px-1.5 py-0.5 rounded shadow-sm">
                        {item.label}
                      </div>
                    )}
                    <button
                      onClick={() => onDeleteFromTemp(item.id)}
                      className="absolute -top-2 -right-2 bg-white rounded-full p-1 shadow-md text-rose-500 hover:bg-rose-50 opacity-0 group-hover:opacity-100 transition-all border border-slate-100"
                    >
                      <X size={12} strokeWidth={3} />
                    </button>
                  </div>
                )
              })}
            </div>
          )}
        </div>
      </div>

      <div
        className="absolute right-0 top-0 bottom-0 w-1 cursor-ew-resize hover:bg-indigo-400 transition-colors z-30"
        onMouseDown={() => setResizing(true)}
      />
    </div>
  );
});

export default function App() {
  // 初期表示のAppIDを決定（URLパラメータ > LocalStorage > Default）
  const getInitialAppId = () => {
    const params = new URLSearchParams(window.location.search);
    const urlProject = params.get('project');
    if (urlProject) return urlProject;

    const saved = localStorage.getItem('daiwari_active_workspace');
    return saved || DEFAULT_APP_ID;
  };

  const [appId, setAppId] = useState(getInitialAppId());
  const [firebaseUser, setFirebaseUser] = useState(null);
  const [isAuthenticated, setIsAuthenticated] = useState(false);

  // 認証情報の永続化チェック
  useEffect(() => {
    const savedAuth = localStorage.getItem('daiwari_is_authenticated');
    if (savedAuth === 'true') {
      setIsAuthenticated(true);
      if (!USE_LOCAL_STORAGE && auth) {
        signInAnonymously(auth).catch((error) => {
          console.error("Auto sign-in failed", error);
        });
      }
    }
  }, []);

  // Data State
  const [sheets, setSheets] = useState([]);
  const [images, setImages] = useState([]);
  const [tempItems, setTempItems] = useState([]);
  const [excludedItems, setExcludedItems] = useState([]);
  const [salesData, setSalesData] = useState(null); // { code: [{name, spec, count}] }
  const [salesDataLastUpdated, setSalesDataLastUpdated] = useState(null);

  // UI State
  const [viewMode, setViewMode] = useState('list');
  const [zoomScale, setZoomScale] = useState(1);
  const [activeSheetId, setActiveSheetId] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
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
  const [isQuickHelpMode, setIsQuickHelpMode] = useState(false);
  const [quickHelpPopup, setQuickHelpPopup] = useState(null);
  const logoTapCountRef = useRef(0);
  const logoTapTimeoutRef = useRef(null);
  const quickHelpHighlightRef = useRef(null);
  const [isSalesMode, setIsSalesMode] = useState(false); // 実績モード
  const [isLabelSelectionMode, setIsLabelSelectionMode] = useState(false);

  // Sales Popup State
  const [hoveredSalesData, setHoveredSalesData] = useState(null);
  const [salesPopupPos, setSalesPopupPos] = useState(null);
  const closeTimeoutRef = useRef(null);

  const [isPageSelectionMode, setIsPageSelectionMode] = useState(false);
  const [selectedSheetIds, setSelectedSheetIds] = useState(new Set());

  const [isProcessing, setIsProcessing] = useState(false);
  const [progressValue, setProgressValue] = useState(0);
  const [progressMax, setProgressMax] = useState(100);
  const [progressMessage, setProgressMessage] = useState("");
  const [isDataLoaded, setIsDataLoaded] = useState(false); // データ読み込み完了フラグ

  // コレクション参照を appId に依存させる
  const sheetsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'sheets'), [appId, db]);
  const imagesCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'images'), [appId, db]);
  const tempShelfCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'tempShelf'), [appId, db]);
  const excludedItemsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'excludedItems'), [appId, db]);
  const settingsCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'settings'), [appId, db]);
  const salesChunksCollection = useMemo(() => USE_LOCAL_STORAGE ? null : collection(db, 'artifacts', appId, 'public', 'data', 'salesDataChunks'), [appId, db]);

  // --- Auth & Init ---
  useEffect(() => {
    // ローカルストレージモードの場合はFirebase認証をスキップ
    if (USE_LOCAL_STORAGE) {
      setFirebaseUser({ uid: 'local_user' }); // ダミーユーザー

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

          setSheets(savedSheets || []);
          setImages(normalizedSavedImages);
          setTempItems(savedTempItems || []);
          setExcludedItems(savedExcludedItems || []);
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
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          if (!auth.currentUser) {
            await signInAnonymously(auth);
          }
        }
      } catch (error) {
        console.error("Auth initialization failed:", error);
      }
    };
    initialize();

    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setFirebaseUser(u);
    });
    return () => unsubscribe();
  }, []);

  const handleAuthenticated = (workspaceId) => {
    if (workspaceId) {
      setAppId(workspaceId);
      localStorage.setItem('daiwari_active_workspace', workspaceId);
      // URLを更新して共有しやすくする
      const url = new URL(window.location);
      url.searchParams.set('project', workspaceId);
      window.history.replaceState({}, '', url);
    }
    setIsAuthenticated(true);
    localStorage.setItem('daiwari_is_authenticated', 'true');
  };

  // --- Data Sync ---
  // 自動保存 (Auto-Save) - IndexedDB with Debounce
  const saveTimeoutRef = useRef(null);

  useEffect(() => {
    if (USE_LOCAL_STORAGE && isDataLoaded) {
      // 既存のタイマーをクリア
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current);
      }

      // 500ms のデバウンス処理で連続操作時のI/O負荷を削減
      saveTimeoutRef.current = setTimeout(async () => {
        try {
          await Promise.all([
            idbHelper.setItem('sheets', sheets),
            idbHelper.setItem('images', images),
            idbHelper.setItem('tempItems', tempItems),
            idbHelper.setItem('excludedItems', excludedItems),
            salesData ? idbHelper.setItem('salesData', salesData) : Promise.resolve()
          ]);
        } catch (err) {
          console.error("Auto-save failed:", err);
        }
      }, 500);
    }

    // クリーンアップ関数
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
      setImages(normalizedImages);
    }
  }, [images, isDataLoaded]);

  useEffect(() => {
    // 詳細単一表示以外ではラベル配置モードを自動解除
    if (viewMode !== 'single' && isLabelSelectionMode) {
      setIsLabelSelectionMode(false);
    }
  }, [viewMode, isLabelSelectionMode]);

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

    if (!isAuthenticated || !firebaseUser) return;

    const unsubscribeSheets = onSnapshot(sheetsCollection, (snapshot) => {
      const loadedSheets = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      loadedSheets.sort((a, b) => (a.createdAt?.seconds || 0) - (b.createdAt?.seconds || 0));
      setSheets(loadedSheets);
    }, (err) => console.error("Sheet Sync Error", err));

    const unsubscribeImages = onSnapshot(imagesCollection, (snapshot) => {
      const loadedImages = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const loadedImageDataById = {};
      loadedImages.forEach((img) => {
        if (img?.id && (img?.data || img?.image)) {
          loadedImageDataById[img.id] = img.data || img.image;
        }
      });
      const normalizedLoadedImages = normalizeStockImages(loadedImages, loadedImageDataById);
      setImages(normalizedLoadedImages);
    }, (err) => console.error("Image Sync Error", err));

    const unsubscribeTemp = onSnapshot(tempShelfCollection, (snapshot) => {
      const loadedTemps = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      loadedTemps.sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      setTempItems(loadedTemps);
    }, (err) => console.error("Temp Shelf Sync Error", err));

    const unsubscribeExcluded = onSnapshot(excludedItemsCollection, (snapshot) => {
      const loadedExcluded = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      loadedExcluded.sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      setExcludedItems(loadedExcluded);
    }, (err) => console.error("Excluded Items Sync Error", err));

    const unsubscribeSales = onSnapshot(salesChunksCollection, (snapshot) => {
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
    });

    const unsubscribeMeta = onSnapshot(doc(settingsCollection, 'salesDataMeta'), (docSnap) => {
      if (docSnap.exists()) {
        setSalesDataLastUpdated(docSnap.data().updatedAt?.toDate() || null);
      }
    });

    return () => {
      unsubscribeSheets();
      unsubscribeImages();
      unsubscribeTemp();
      unsubscribeExcluded();
      unsubscribeSales();
      unsubscribeMeta();
    };
  }, [isAuthenticated, firebaseUser, sheetsCollection, imagesCollection, tempShelfCollection, excludedItemsCollection, settingsCollection, salesChunksCollection]);

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

  // --- Common CSV Parser ---
  const parseCSVLine = (line) => {
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  };

  const canPlacePanelAt = (startIdx, rowSpan, colSpan, occupied) => {
    const startRow = Math.floor(startIdx / 4);
    const startCol = startIdx % 4;
    if (startRow + rowSpan > 4 || startCol + colSpan > 4) return false;

    for (let r = 0; r < rowSpan; r++) {
      for (let c = 0; c < colSpan; c++) {
        const idx = (startRow + r) * 4 + (startCol + c);
        if (occupied.has(idx)) return false;
      }
    }
    return true;
  };

  const findFirstPlaceableIndex = (rowSpan, colSpan, occupied, startIndex = 0) => {
    for (let i = startIndex; i < 16; i++) {
      if (canPlacePanelAt(i, rowSpan, colSpan, occupied)) return i;
    }
    // もし指定の開始位置から見つからなければ、最初から探し直す (フォールバック)
    if (startIndex > 0) {
      for (let i = 0; i < startIndex; i++) {
        if (canPlacePanelAt(i, rowSpan, colSpan, occupied)) return i;
      }
    }
    return -1;
  };

  const fillPanelArea = (panels, startIdx, rowSpan, colSpan, occupied) => {
    const startRow = Math.floor(startIdx / 4);
    const startCol = startIdx % 4;

    for (let r = 0; r < rowSpan; r++) {
      for (let c = 0; c < colSpan; c++) {
        const idx = (startRow + r) * 4 + (startCol + c);
        occupied.add(idx);
        if (idx !== startIdx) {
          panels[idx] = { ...panels[idx], hidden: true };
        }
      }
    }
  };

  // --- Sales CSV Import Logic ---
  const handleImportSalesCSV = async (file) => {
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
        await deleteBatch.commit();

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
          await writeBatchChunk.commit();
        }
      } else {
        await batch.commit();
      }

      await setDoc(doc(settingsCollection, 'salesDataMeta'), {
        updatedAt: serverTimestamp(),
        totalItems: entries.length
      });

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
    if (!canMerge) return;
    const { sheetId, indices } = selection;
    const sheet = sheets.find(s => s.id === sheetId);
    if (!sheet) return;

    const coords = indices.map(getCoords);
    const minRow = Math.min(...coords.map(c => c.row));
    const maxRow = Math.max(...coords.map(c => c.row));
    const minCol = Math.min(...coords.map(c => c.col));
    const maxCol = Math.max(...coords.map(c => c.col));
    const rowSpan = maxRow - minRow + 1;
    const colSpan = maxCol - minCol + 1;
    const primaryIndex = minRow * 4 + minCol;
    const sizeType = getSizeType(rowSpan, colSpan);

    const newPanels = [...sheet.panels];
    newPanels[primaryIndex] = { ...newPanels[primaryIndex], rowSpan, colSpan, hidden: false, sizeType };
    indices.forEach(idx => {
      if (idx !== primaryIndex) {
        newPanels[idx] = { ...newPanels[idx], hidden: true, image: null, text: '', rowSpan: 1, colSpan: 1 };
      }
    });

    if (USE_LOCAL_STORAGE) {
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
      await updateDoc(doc(sheetsCollection, sheetId), { panels: newPanels });
    } catch (error) {
      console.error("Merge error:", error);
      showAlert("結合に失敗しました");
    } finally {
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
    }
  }, [canMerge, selection, sheets, sheetsCollection]);

  const handleSplit = useCallback(async () => {
    const { sheetId, indices } = selection;
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

    if (USE_LOCAL_STORAGE) {
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
      await updateDoc(doc(sheetsCollection, sheetId), { panels: newPanels });
    } catch (error) {
      console.error("Split error:", error);
      showAlert("分離に失敗しました");
    } finally {
      setSelection({ sheetId: null, indices: [] });
      setIsMergeMode(false);
    }
  }, [selection, sheets, sheetsCollection]);

  // --- Core Actions ---

  const handleAddSheet = useCallback(async () => {
    // 認証チェック: LocalStorageモードならUI認証のみ、FirebaseモードならFirebase認証も確認
    if (!isAuthenticated) return;

    if (!USE_LOCAL_STORAGE && !auth.currentUser) {
      try {
        console.log("Attempting JIT sign-in...");
        await signInAnonymously(auth);
      } catch (e) {
        console.error("JIT Auth failed", e);
        showAlert("サーバーへの接続に失敗しました。再読み込みしてください。");
        return;
      }
    }

    const newOrder = sheets.length > 0 ? Math.max(...sheets.map(s => s.order || 0)) + 1 : 0;

    // Fix: Create independent objects for each panel to prevent reference sharing
    const defaultPanels = Array(16).fill(null).map(() => ({
      image: null,
      imageId: null,
      text: '',
      label: null,
      code: null,
      rowSpan: 1,
      colSpan: 1,
      hidden: false,
      sizeType: '1/16（1コマ）',
      isText: false
    }));

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
        panels: defaultPanels,
        createdAt: serverTimestamp()
      });
    } catch (e) {
      console.error("Error adding sheet: ", e);
      showAlert("ページの追加に失敗しました。");
    }
  }, [isAuthenticated, sheets, sheetsCollection]);

  const handleUpdatePanel = useCallback(async (sheetId, panelIndex, newData) => {
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

    if (!sheetsCollection || !tempShelfCollection) return;

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      await runTransaction(db, async (transaction) => {
        const sheetSnap = await transaction.get(sheetRef);
        if (!sheetSnap.exists()) return;

        const serverPanels = [...(sheetSnap.data().panels || [])];
        const serverPanel = serverPanels[panelIndex] || {};
        const mergedPanel = { ...serverPanel, ...panelPatch };

        if (isPanelDataEqual(serverPanel, mergedPanel)) return;
        serverPanels[panelIndex] = mergedPanel;
        transaction.update(sheetRef, { panels: serverPanels });
      });
    } catch (error) {
      console.error("Panel update transaction failed:", error);
      showAlert("同時編集との競合が発生しました。再度お試しください。");
    }
  }, [sheets, sheetsCollection, showAlert]);

  // --- Temp & Excluded Logic (Restored) ---

  const handleMoveToTemp = async (sheetId, panelIndex, movedText) => {
    if (USE_LOCAL_STORAGE) {
      const sheet = sheets.find(s => s.id === sheetId);
      if (!sheet) return;
      const panel = sheet.panels[panelIndex];
      if (!hasPanelTransferableContent(panel)) return;

      const newTempItem = {
        id: idbHelper.generateId(),
        image: panel.image || null,
        imageId: panel.imageId || null,
        label: panel.label || null,
        code: panel.code || null,
        text: movedText !== undefined ? movedText : (panel.text || ''),
        isText: panel.isText || false,
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

    if (!sheetsCollection || !excludedItemsCollection) return;

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      const tempRef = doc(tempShelfCollection);

      await runTransaction(db, async (transaction) => {
        const sheetSnap = await transaction.get(sheetRef);
        if (!sheetSnap.exists()) return;

        const serverPanels = [...(sheetSnap.data().panels || [])];
        const sourcePanel = serverPanels[panelIndex] || {};
        if (!hasPanelTransferableContent(sourcePanel)) return;

        transaction.set(tempRef, {
          image: sourcePanel.image || null,
          imageId: sourcePanel.imageId || null,
          label: sourcePanel.label || null,
          code: sourcePanel.code || null,
          text: movedText !== undefined ? movedText : (sourcePanel.text || ''),
          isText: sourcePanel.isText || false,
          originalName: "退避アイテム",
          createdAt: serverTimestamp()
        });

        serverPanels[panelIndex] = clearPanelTransferableContent(sourcePanel);
        transaction.update(sheetRef, { panels: serverPanels });
      });
    } catch (error) {
      console.error("Move to temp transaction failed:", error);
      showAlert("仮置き場への移動に失敗しました。再度お試しください。");
    }
  };

  const handleDeleteFromTemp = async (id) => {
    if (USE_LOCAL_STORAGE) {
      const newTempItems = tempItems.filter(item => item.id !== id);
      setTempItems(newTempItems);
      // localStorageHelper.setItem('tempItems', newTempItems); // Auto-save handles this
      return;
    }
    await deleteDoc(doc(tempShelfCollection, id));
  };

  const handleMoveToExcluded = async (sheetId, panelIndex, movedText) => {
    if (USE_LOCAL_STORAGE) {
      const sheet = sheets.find(s => s.id === sheetId);
      if (!sheet) return;
      const panel = sheet.panels[panelIndex];
      if (!hasPanelTransferableContent(panel)) return;

      const newExcludedItem = {
        id: idbHelper.generateId(),
        image: panel.image || null,
        imageId: panel.imageId || null,
        label: panel.label || null,
        code: panel.code || null,
        text: movedText !== undefined ? movedText : (panel.text || ''),
        isText: panel.isText || false,
        originalName: "掲載除外",
        createdAt: { seconds: Date.now() / 1000 }
      };

      const newExcludedItems = [newExcludedItem, ...excludedItems];
      setExcludedItems(newExcludedItems);
      // localStorageHelper.setItem('excludedItems', newExcludedItems); // Auto-save handles this

      const updatedPanels = [...sheet.panels];
      updatedPanels[panelIndex] = clearPanelTransferableContent(updatedPanels[panelIndex]);

      const newSheets = sheets.map(s => s.id === sheetId ? { ...s, panels: updatedPanels } : s);
      setSheets(newSheets);
      // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      return;
    }

    if (!sheetsCollection) return;

    try {
      const sheetRef = doc(sheetsCollection, sheetId);
      const excludedRef = doc(excludedItemsCollection);

      await runTransaction(db, async (transaction) => {
        const sheetSnap = await transaction.get(sheetRef);
        if (!sheetSnap.exists()) return;

        const serverPanels = [...(sheetSnap.data().panels || [])];
        const sourcePanel = serverPanels[panelIndex] || {};
        if (!hasPanelTransferableContent(sourcePanel)) return;

        transaction.set(excludedRef, {
          image: sourcePanel.image || null,
          imageId: sourcePanel.imageId || null,
          label: sourcePanel.label || null,
          code: sourcePanel.code || null,
          text: movedText !== undefined ? movedText : (sourcePanel.text || ''),
          isText: sourcePanel.isText || false,
          originalName: "掲載除外",
          createdAt: serverTimestamp()
        });

        serverPanels[panelIndex] = clearPanelTransferableContent(sourcePanel);
        transaction.update(sheetRef, { panels: serverPanels });
      });
    } catch (error) {
      console.error("Move to excluded transaction failed:", error);
      showAlert("除外リストへの移動に失敗しました。再度お試しください。");
    }
  };

  const handleDeleteFromExcluded = async (id) => {
    requestConfirm(
      "掲載除外リストから完全に削除しますか？\n（復元できません）",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newExcludedItems = excludedItems.filter(item => item.id !== id);
          setExcludedItems(newExcludedItems);
          // localStorageHelper.setItem('excludedItems', newExcludedItems); // Auto-save handles this
          return;
        }
        await deleteDoc(doc(excludedItemsCollection, id));
      }
    );
  };

  // 除外リスト一括削除機能
  const handleBulkDeleteExcluded = async () => {
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
          await batch.commit();
          showAlert("除外リストを空にしました。");
        } catch (err) {
          console.error("Bulk delete excluded failed", err);
          showAlert("一括削除に失敗しました。");
        }
      }
    );
  };

  const handleExportExcludedCSV = () => {
    const headers = ['介援隊コード', '画像名', 'ラベル', '登録日時'];
    const rows = excludedItems.map(item => {
      const date = item.createdAt?.toDate
        ? item.createdAt.toDate().toLocaleString()
        : (item.createdAt?.seconds ? new Date(item.createdAt.seconds * 1000).toLocaleString() : new Date().toLocaleString());

      return [
        item.code || '',
        item.originalName || '',
        item.label || '',
        date
      ].join(',');
    });

    const csvContent = "\uFEFF" + [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `excluded_items_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handlePanelUpdateWithCheck = (sheetId, panelIndex, newData) => {
    const sanitizedData = { ...newData };

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
        deleteDoc(doc(excludedItemsCollection, sanitizedData.fromExcludedId));
      }
      delete sanitizedData.fromExcludedId;
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

  const handleMoveToStock = useCallback(async (sheetId, panelIndex) => {
    const sheet = sheets.find(s => s.id === sheetId);
    if (!sheet) return;
    const panel = sheet.panels[panelIndex];
    if (!panel) return;
    handlePanelUpdateWithCheck(sheetId, panelIndex, {
      ...panel,
      image: null, imageId: null, label: null, code: null, text: '', isText: false
    });
  }, [sheets]);

  const handleMovePanel = async (fromSheetId, fromIndex, toSheetId, toIndex, movedText) => {
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

        newPanels[toIndex] = {
          ...newPanels[toIndex],
          image: dataToMove.image || null,
          imageId: dataToMove.imageId || null,
          label: dataToMove.label || null,
          code: dataToMove.code || null,
          text: movedText !== undefined ? movedText : (dataToMove.text || ''),
          isText: !!dataToMove.isText
        };

        const newSheets = sheets.map(s => s.id === fromSheetId ? { ...s, panels: newPanels } : s);
        setSheets(newSheets);
        // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
      } else {
        const newFromPanels = [...fromSheet.panels];
        const dataToMove = { ...newFromPanels[fromIndex] };

        newFromPanels[fromIndex] = clearPanelTransferableContent(newFromPanels[fromIndex]);

        const newToPanels = [...toSheet.panels];
        newToPanels[toIndex] = {
          ...newToPanels[toIndex],
          image: dataToMove.image || null,
          imageId: dataToMove.imageId || null,
          label: dataToMove.label || null,
          code: dataToMove.code || null,
          text: movedText !== undefined ? movedText : (dataToMove.text || ''),
          isText: !!dataToMove.isText
        };

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

      await runTransaction(db, async (transaction) => {
        const fromSnap = await transaction.get(fromSheetRef);
        if (!fromSnap.exists()) return;

        const fromPanels = [...(fromSnap.data().panels || [])];
        const dataToMove = { ...(fromPanels[fromIndex] || {}) };
        if (!hasPanelTransferableContent(dataToMove)) return;

        if (fromSheetId === toSheetId) {
          fromPanels[fromIndex] = clearPanelTransferableContent(dataToMove);
          const targetPanel = fromPanels[toIndex] || {};
          fromPanels[toIndex] = {
            ...targetPanel,
            image: dataToMove.image || null,
            imageId: dataToMove.imageId || null,
            label: dataToMove.label || null,
            code: dataToMove.code || null,
            text: movedText !== undefined ? movedText : (dataToMove.text || ''),
            isText: !!dataToMove.isText
          };
          transaction.update(fromSheetRef, { panels: fromPanels });
          return;
        }

        const toSnap = await transaction.get(toSheetRef);
        if (!toSnap.exists()) return;

        const toPanels = [...(toSnap.data().panels || [])];
        fromPanels[fromIndex] = clearPanelTransferableContent(dataToMove);
        const targetPanel = toPanels[toIndex] || {};
        toPanels[toIndex] = {
          ...targetPanel,
          image: dataToMove.image || null,
          imageId: dataToMove.imageId || null,
          label: dataToMove.label || null,
          code: dataToMove.code || null,
          text: movedText !== undefined ? movedText : (dataToMove.text || ''),
          isText: !!dataToMove.isText
        };

        transaction.update(fromSheetRef, { panels: fromPanels });
        transaction.update(toSheetRef, { panels: toPanels });
      });
    } catch (error) {
      console.error("Move panel transaction failed:", error);
      showAlert("同時編集との競合が発生しました。再度お試しください。");
    }
  };

  // --- Image & Bulk Actions ---

  const handleUploadImage = async (e) => {
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
          createdAt: { seconds: Date.now() / 1000 }
        };

        if (USE_LOCAL_STORAGE) {
          newImages.push(newImage);
        } else {
          await addDoc(imagesCollection, {
            name: file.name,
            data: compressedDataUrl,
            createdAt: serverTimestamp()
          });
        }
        successCount++;
      } catch (err) {
        console.error(`Failed to upload ${file.name}:`, err);
        failCount++;
      }
    });

    try {
      await Promise.all(uploadPromises);

      if (USE_LOCAL_STORAGE && newImages.length > 0) {
        const updatedImages = [...images, ...newImages];
        setImages(updatedImages);
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
        if (!imgId) {
          showAlert("削除対象の画像IDが見つかりませんでした。");
          return;
        }
        await deleteDoc(doc(imagesCollection, imgId));
      }
    );
  };

  const handleBulkDeleteImages = async (imageTargets) => {
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
            await batch.commit();
          }
        } catch (err) {
          console.error("Bulk image delete failed", err);
          showAlert("一括削除に失敗しました");
        }
      }
    );
  };

  // --- Page Actions ---

  const handleNavigatePage = (direction) => {
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
    if (USE_LOCAL_STORAGE) {
      const newSheets = sheets.map(s => s.id === sheetId ? { ...s, genre: newGenre } : s);
      setSheets(newSheets);
      return;
    }
    const sheetRef = doc(sheetsCollection, sheetId);
    await updateDoc(sheetRef, { genre: newGenre });
  };

  const handleDeleteSheet = (sheetId) => {
    requestConfirm(
      "このページを削除しますか？\n（この操作は取り消せません）",
      async () => {
        if (USE_LOCAL_STORAGE) {
          const newSheets = sheets.filter(s => s.id !== sheetId);
          setSheets(newSheets);
          if (activeSheetId === sheetId) setActiveSheetId(null);
          return;
        }
        await deleteDoc(doc(sheetsCollection, sheetId));
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

  const handleSwapPages = async () => {
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

        const batch = writeBatch(db);
        const ref1 = doc(sheetsCollection, id1);
        const ref2 = doc(sheetsCollection, id2);

        batch.update(ref1, {
          genre: sheet2.genre,
          panels: sheet2.panels
        });
        batch.update(ref2, {
          genre: sheet1.genre,
          panels: sheet1.panels
        });

        try {
          await batch.commit();
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
    if (selectedSheetIds.size === 0) return;

    requestConfirm(
      `${selectedSheetIds.size}ページ分の画像を外しますか？\n（画像は未配置リストに戻ります）`,
      async () => {
        // 回収する画像を収集
        const recoveredImages = [];
        selectedSheetIds.forEach(id => {
          const sheet = sheets.find(s => s.id === id);
          if (sheet) {
            sheet.panels.forEach(p => {
              const resolvedImage = p.image || (p.imageId ? imageDataById?.[p.imageId] : null);
              if (resolvedImage) {
                const recoveredId = idbHelper.generateId();
                recoveredImages.push({
                  id: recoveredId,
                  name: p.code ? `${p.code}.png` : `recovered-${recoveredId}.png`,
                  data: resolvedImage,
                  createdAt: { seconds: Date.now() / 1000 }
                });
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
                  isText: false
                }))
              };
            }
            return s;
          });

          setImages(prev => normalizeStockImages([...prev, ...recoveredImages]));
          setSheets(newSheets);
          // localStorageHelper.setItem('sheets', newSheets); // Auto-save handles this
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
          return;
        }

        const batch = writeBatch(db);

        // Recovered images to Firestore
        recoveredImages.forEach(img => {
          const ref = doc(imagesCollection);
          batch.set(ref, {
            name: img.name,
            data: img.data,
            createdAt: serverTimestamp()
          });
        });

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
            isText: false
          }));

          const ref = doc(sheetsCollection, id);
          batch.update(ref, { panels: newPanels });
        });

        try {
          await batch.commit();
          setSelectedSheetIds(new Set());
          setIsPageSelectionMode(false);
          showAlert("画像を解除し、未配置リストに戻しました");
        } catch (err) {
          console.error("Bulk clear images failed", err);
          showAlert("一括解除に失敗しました");
        }
      }
    );
  };

  const handleBulkDelete = async () => {
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
          await batch.commit();
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
      const headers = ['ジャンル', 'ページ数', '追番', 'コマ番号', '介援隊コード', 'コマ数', '', 'テキスト情報', '座標', 'コマID'];
      const rows = [];

      sheets.forEach((sheet, sheetIndex) => {
        const genreLabel = GENRES.find(g => g.id === sheet.genre)?.label || '未設定';
        const pageNum = sheetIndex + 1;
        let visibleCounter = 0;
        let frameCounter = 0;

        sheet.panels.forEach((panel, panelIndex) => {
          if (panel.hidden) return;
          frameCounter++;
          const isSpecialDummy = panel.label === '埋草' || panel.label === 'タイトル';
          let panelNum = '';
          if (!isSpecialDummy) {
            visibleCounter++;
            panelNum = visibleCounter;
          }
          let codeVal = panel.code || '';
          if (panel.label) {
            codeVal = 'ダミーコマ';
          }
          let sizeVal = panel.sizeType || getSizeType(panel.rowSpan || 1, panel.colSpan || 1);
          let textVal = panel.text || '';
          if (/[,"\n]/.test(textVal)) {
            textVal = `"${textVal.replace(/"/g, '""')}"`;
          }

          // I列: グリッド座標 (X1Y1 〜 X4Y4)
          const gridRow = Math.floor(panelIndex / 4) + 1; // 1始まり
          const gridCol = (panelIndex % 4) + 1;           // 1始まり
          const coordVal = `X${gridCol}Y${gridRow}`;

          // J列: コマID（パネルデータに保持している値を出力）
          const panelIdVal = panel.panelId || '';

          rows.push([
            genreLabel,
            pageNum,
            panelNum,
            frameCounter,
            codeVal,
            sizeVal,
            '',
            textVal,
            coordVal,
            panelIdVal
          ].join(','));
        });
      });

      const csvContent = "\uFEFF" + [headers.join(','), ...rows].join('\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `daiwari_export_${new Date().toISOString().slice(0, 10)}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error("Export failed", err);
      showAlert("CSV出力に失敗しました: " + err.message);
    }
  };

  // --- CSV Import Logic (for Pages) ---
  const handleImportCSV = async (e) => {
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

      const searchableImages = images.map(img => ({
        id: img.id,
        name: img.name || '',
        lowerName: (img.name || '').toLowerCase(),
        data: img.data
      }));

      const sheetUpdates = {};
      let maxPageIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row.trim()) continue;

        if (i % 50 === 0) {
          setProgressValue(Math.floor((i / rows.length) * 100));
          setProgressMessage(`解析中... (${i}/${rows.length}行)`);
          await new Promise(resolve => setTimeout(resolve, 0));
        }

        const cols = parseCSVLine(row);
        if (cols.length < 6) continue;

        // カラム定義: 0:ジャンル, 1:ページ数, 2:追番, 3:コマ番号, 4:介援隊コード, 5:コマ数, 6:ダミーラベル種別, 7:テキスト情報, 8:座標(無視), 9:コマID
        const genreLabel = cols[0];
        const pageNum = parseInt(cols[1], 10);
        const panelNumRaw = parseInt(cols[2], 10);
        const frameNumRaw = parseInt(cols[3], 10);
        const frameNum = (!isNaN(frameNumRaw) && frameNumRaw > 0)
          ? frameNumRaw
          : ((!isNaN(panelNumRaw) && panelNumRaw > 0) ? panelNumRaw : NaN);
        const panelNum = (!isNaN(panelNumRaw) && panelNumRaw > 0)
          ? panelNumRaw
          : ((!isNaN(frameNumRaw) && frameNumRaw > 0) ? frameNumRaw : NaN);
        const codeVal = cols[4] === 'ダミーコマ' ? '' : (cols[4] || '').trim();
        const isDummyMarker = cols[4] === 'ダミーコマ';
        const sizeVal = cols[5] || '1/16（1コマ）';
        const textVal = (cols[7] || '').trim();
        // J列: コマID（台割には反映しない。介援隊コードに紐づけてデータとして保持）
        const panelIdVal = (cols[9] || '').trim();

        const isFixed = !isNaN(frameNum) && frameNum > 0;

        // ページ番号は必須。コマ番号か追番のどちらかは必須。
        if (isNaN(pageNum)) continue;
        if (!isFixed && (isNaN(panelNum))) continue;

        const pageIndex = pageNum - 1;
        if (pageIndex > maxPageIndex) maxPageIndex = pageIndex;

        if (!sheetUpdates[pageIndex]) {
          sheetUpdates[pageIndex] = { genre: null, contentItems: [] };
        }

        const genreObj = GENRES.find(g => g.label === genreLabel);
        if (genreObj) sheetUpdates[pageIndex].genre = genreObj.id;

        let targetImageId = null;
        let targetCode = null;
        let targetLabel = null;
        let isText = false;

        // 優先順位: 1.テキスト 2.ダミーコマ 3.コード(画像)
        if (textVal) {
          targetLabel = 'テキスト';
          isText = true;
        } else if (isDummyMarker) {
          // ダミーコマの種別を判定: cols[6]（ラベル列）、cols[5]（コマ数列）、cols[4]の順にチェック
          const dummyLabel = (cols[6] || '').trim();
          if (dummyLabel === 'タイトル' || sizeVal.includes('タイトル')) {
            targetLabel = 'タイトル';
          } else if (dummyLabel === '埋草' || sizeVal.includes('埋草')) {
            targetLabel = '埋草';
          } else if (dummyLabel === '新規商品' || dummyLabel === '新規商品未確定') {
            targetLabel = '新規商品未確定';
          } else if (dummyLabel) {
            // CSVに記載のあるラベルをそのまま使用
            targetLabel = dummyLabel;
          } else {
            targetLabel = '新規商品未確定';
          }
        } else if (codeVal) {
          const normalizedToken = codeVal.normalize('NFKC').replace(/[-\s]/g, '').toUpperCase();
          const isLikelyProductCode = /^[A-Z]{1,2}\d{3,5}$/.test(normalizedToken);

          if (!isLikelyProductCode) {
            targetLabel = codeVal;
          } else {
            targetCode = normalizedToken;

            // Helper: 全角英数字を半角に変換＆小文字化
            const normalizeStr = (s) => {
              return s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).toLowerCase();
            };

            const searchCode = normalizeStr(targetCode);
            const isNumericSearch = /^\d+$/.test(searchCode);
            const searchNum = isNumericSearch ? parseInt(searchCode, 10) : null;

            // ベストマッチ検索 (スコアリング方式)
            let bestMatchImg = null;
            let bestScore = 0;

            for (const img of searchableImages) {
              const normImgName = normalizeStr(img.name);
              const stem = normImgName.lastIndexOf('.') !== -1
                ? normImgName.substring(0, normImgName.lastIndexOf('.'))
                : normImgName;

              let score = 0;

              // 1. 完全一致 (Score: 100) - 拡張子なし
              if (stem === searchCode) {
                score = 100;
              }
              // 2. 数値トークン完全一致 (Score: 80)
              else if (isNumericSearch) {
                const numTokens = stem.match(/\d+/g);
                if (numTokens) {
                  if (numTokens.some(t => parseInt(t, 10) === searchNum)) {
                    if (numTokens.length === 1 && stem.replace(/\d+/g, '').length < stem.length) {
                      score = 80;
                    } else {
                      score = 60;
                    }
                  }
                }
              }
              // 3. 記号除去しての一致 (Score: 50)
              else {
                const clean = (s) => s.replace(/[^a-z0-9]/g, '');
                if (clean(stem) === clean(searchCode)) {
                  score = 50;
                }
                // 4. 包含一致 (Score: 30) - 非数値のみ
                else if (stem.includes(searchCode)) {
                  score = 30;
                }
              }

              if (score > bestScore) {
                bestScore = score;
                bestMatchImg = img;
                if (score === 100) break; // 完全一致なら即決
              }
            }

            // スコア90以上のマッチングのみを有効とする
            const MATCH_THRESHOLD = 90;
            if (bestMatchImg && bestScore >= MATCH_THRESHOLD) {
              targetImageId = bestMatchImg.id || null;
              // マッチング詳細を記録
              sheetUpdates[pageIndex].matchDetails = sheetUpdates[pageIndex].matchDetails || [];
              sheetUpdates[pageIndex].matchDetails.push({
                code: targetCode,
                imageName: bestMatchImg.name,
                score: bestScore,
                csvRow: i + 1
              });
            } else if (targetCode) {
              // マッチしなかったコードを記録
              sheetUpdates[pageIndex].unmatchedCodes = sheetUpdates[pageIndex].unmatchedCodes || [];
              sheetUpdates[pageIndex].unmatchedCodes.push({
                code: targetCode,
                csvRow: i + 1,
                bestScore: bestScore,
                bestMatch: bestMatchImg ? bestMatchImg.name : 'なし'
              });
            }
          }
        }

        sheetUpdates[pageIndex].contentItems.push({
          isFixed: isFixed,
          frameNo: isFixed ? frameNum : -1,
          order: isNaN(panelNum) ? 9999 : panelNum,
          data: {
            code: targetCode,
            image: null,
            imageId: targetImageId,
            label: targetLabel,
            sizeType: sizeVal,
            text: textVal,
            isText: isText,
            panelId: panelIdVal || null  // J列から読み込んだコマID（台割には非表示）
          }
        });
      }

      // ページを不足分作成
      const finalPageCount = maxPageIndex + 1;
      let localSheets = [...sheets];

      // 足りないページをパディング
      while (localSheets.length < finalPageCount) {
        localSheets.push({
          id: idbHelper.generateId(),
          createdAt: { seconds: Date.now() / 1000 },
          genre: 'none',
          panels: Array(16).fill(null).map(() => ({
            image: null, imageId: null, text: '', label: null, code: null,
            rowSpan: 1, colSpan: 1, hidden: false, sizeType: '1/16（1コマ）'
          }))
        });
      }

      const importSummary = {
        total: 0,
        fixedSuccess: 0,
        fixedFailed: 0,
        autoSuccess: 0,
        autoFailed: 0,
        details: [],
        matchedImages: [], // マッチした画像の詳細
        notMatchedCodes: [], // マッチしなかったコードのリスト
        imageUsageCount: {} // 画像の使用回数
      };

      setProgressMax(finalPageCount);
      for (let i = 0; i < finalPageCount; i++) {
        const update = sheetUpdates[i];
        if (!update) continue;

        setProgressValue(i + 1);
        setProgressMessage(`${i + 1}ページ目を再構成中...`);
        await new Promise(resolve => setTimeout(resolve, 0));

        const currentSheet = { ...localSheets[i] };
        if (update.genre) currentSheet.genre = update.genre;

        const newPanels = Array(16).fill(null).map(() => ({
          image: null, imageId: null, text: '', label: null, code: null,
          rowSpan: 1, colSpan: 1, hidden: false, sizeType: '1/16（1コマ）'
        }));

        const occupied = new Set();

        // === コマ番号順 → 先頭空きスロットへ順次配置 ===
        // コマ番号順にソートし、各コマを「前のコマ配置後の最初の空き位置」に配置する。
        // 例: コマ1(1/8横 2コマ) → idx=0(X1Y1),1(X2Y1)占有
        //     コマ2(1/8横 2コマ) → 次の空き=idx=2(X3Y1),3(X4Y1)占有
        //     コマ3              → 次の空き=idx=4(X1Y2)から
        const allItems = [...update.contentItems].sort((a, b) => {
          const aKey = a.frameNo > 0 ? a.frameNo : (a.order > 0 ? a.order : Number.MAX_SAFE_INTEGER);
          const bKey = b.frameNo > 0 ? b.frameNo : (b.order > 0 ? b.order : Number.MAX_SAFE_INTEGER);
          return aKey - bKey;
        });

        for (const item of allItems) {
          importSummary.total++;
          const { data } = item;
          const { r: rowSpan, c: colSpan } = getSpansFromSizeTypeRobust(data.sizeType);

          // 前のコマ配置後の occupied を考慮し、先頭から最初に配置可能な位置を探す
          // コマ番号 (item.order) を開始座標のヒントとして使用
          // ユーザーの要件: コマ番号1→X1Y1, 2→X2Y1 ... 等。
          // findFirstPlaceableIndex に第4引数として (order-1) を渡す。
          const startCandidate = (item.order > 0 && item.order <= 16) ? item.order - 1 : 0;
          const resolvedStartIdx = findFirstPlaceableIndex(rowSpan, colSpan, occupied, startCandidate);

          if (resolvedStartIdx === -1) {
            importSummary.autoFailed++;
            importSummary.details.push(`・ページ ${i + 1}: 「${data.code || data.label || data.text || '不明'}」（${data.sizeType}）を配置できませんでした。`);
            continue;
          }

          newPanels[resolvedStartIdx] = {
            ...newPanels[resolvedStartIdx],
            ...data,
            rowSpan,
            colSpan,
            hidden: false
          };
          fillPanelArea(newPanels, resolvedStartIdx, rowSpan, colSpan, occupied);
          importSummary.autoSuccess++;
        }

        currentSheet.panels = newPanels;
        localSheets[i] = currentSheet;

        // マッチング詳細をサマリーに集約
        if (update.matchDetails) {
          update.matchDetails.forEach(detail => {
            importSummary.matchedImages.push({
              page: i + 1,
              ...detail
            });
            // 画像使用回数をカウント
            const imgKey = detail.imageName;
            importSummary.imageUsageCount[imgKey] = (importSummary.imageUsageCount[imgKey] || 0) + 1;
          });
        }

        // 未マッチコードをサマリーに集約
        if (update.unmatchedCodes) {
          update.unmatchedCodes.forEach(unmatched => {
            importSummary.notMatchedCodes.push({
              page: i + 1,
              ...unmatched
            });
          });
        }
      }

      let saveFailed = false;
      let saveErrorMessage = '';

      if (!USE_LOCAL_STORAGE) {
        for (let i = 0; i < localSheets.length; i++) {
          const sheet = localSheets[i];
          if (!sheet?.id || !String(sheet.id).startsWith('local_')) continue;

          const newRef = await addDoc(sheetsCollection, {
            genre: sheet.genre || 'none',
            order: sheet.order ?? i,
            panels: sheet.panels || [],
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
              await batch.commit();
              batch = writeBatch(db);
              opCount = 0;
            }

            batch.update(doc(sheetsCollection, targetSheet.id), {
              genre: targetSheet.genre || 'none',
              order: targetSheet.order ?? i,
              panels: targetSheet.panels
            });
            opCount++;
          }
          if (opCount > 0) {
            await batch.commit();
          }
        } catch (err) {
          console.error("Cloud save failed:", err);
          saveFailed = true;
          saveErrorMessage = err.message;
        }
      }

      setIsProcessing(false);

      // 重複使用されている画像を抽出
      const duplicateImages = Object.entries(importSummary.imageUsageCount)
        .filter(([_, count]) => count > 1)
        .map(([name, count]) => `${name} (${count}回)`);

      const report = [
        `取り込みが完了しました。(全${importSummary.total}件)`,
        `・指定通りの配置: ${importSummary.fixedSuccess}件`,
        `・空きへの自動配置: ${importSummary.autoSuccess}件`,
        importSummary.fixedFailed > 0 ? `・指定位置が重複または不足により移動: ${importSummary.fixedFailed}件` : '',
        importSummary.autoFailed > 0 ? `・スペース不足で配置失敗: ${importSummary.autoFailed}件` : '',
        '',
        `【画像マッチング結果】`,
        `・マッチング成功: ${importSummary.matchedImages.length}件`,
        `・マッチング失敗: ${importSummary.notMatchedCodes.length}件`,
        importSummary.notMatchedCodes.length > 0 ? `\n【マッチしなかったコード】` : '',
        ...importSummary.notMatchedCodes.slice(0, 15).map(item =>
          `・P${item.page} 行${item.csvRow}: ${item.code} (ベストマッチ: ${item.bestMatch}, スコア: ${item.bestScore})`
        ),
        importSummary.notMatchedCodes.length > 15 ? `...他 ${importSummary.notMatchedCodes.length - 15} 件` : '',
        duplicateImages.length > 0 ? `\n【重複使用されている画像】` : '',
        ...duplicateImages.slice(0, 10).map(item => `・${item}`),
        duplicateImages.length > 10 ? `...他 ${duplicateImages.length - 10} 件` : '',
        importSummary.details.length > 0 ? "\n【未配置の項目】\n" + importSummary.details.slice(0, 10).join('\n') + (importSummary.details.length > 10 ? '\n...他' : '') : ''
      ].filter(Boolean).join('\n');

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

  const imageDataById = useMemo(() => {
    const map = {};
    images.forEach((img) => {
      if (img?.id && img?.data) {
        map[img.id] = img.data;
      }
    });
    return map;
  }, [images]);

  const currentList = useMemo(() => {
    if (viewMode === 'single') return sheets;
    return genreFilter === 'all' ? sheets : sheets.filter(s => s.genre === genreFilter);
  }, [sheets, genreFilter, viewMode]);

  const currentIndex = useMemo(() => {
    return currentList.findIndex(s => s.id === activeSheetId);
  }, [currentList, activeSheetId]);

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
          await updateDoc(doc(sheetsCollection, activeSheetId), { panels: updatedPanels });
        } catch (err) {
          console.error("Bulk label delete failed", err);
          showAlert("ラベル一括削除に失敗しました。");
        }
      }
    );
  }, [viewMode, activeSheetId, sheets, activeSheetLabelCount, sheetsCollection, requestConfirm, showAlert]);

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
  if (!isAuthenticated) return <AuthGate onAuthenticated={handleAuthenticated} defaultAppId={appId} />;

  return (
    <div className={`flex flex-col h-screen overflow-hidden transition-all duration-700 ease-in-out`} style={{ background: 'var(--app-bg)', color: 'var(--m3-on-surface)' }}>

      {/* Top Navigation Bar - M3 Expressive Style */}
      <div className="h-20 flex items-center justify-between px-6 z-30 flex-shrink-0 relative transition-all" style={{ background: 'var(--m3-surface)', color: 'var(--m3-on-surface)' }}>
        <div className="flex items-center gap-4 flex-shrink-0">
          <div className="flex items-center">
            <div
              className="p-0.5 bg-white shadow-sm cursor-pointer select-none"
              style={{ borderRadius: 'var(--m3-shape-corner-md)' }}
              onClick={handleLogoSecretTap}
              title="台"
            >
              <img
                src="/logo.jpg"
                alt="台割君"
                className="h-16 w-16 object-contain transition-transform hover:scale-105"
                style={{ borderRadius: 'calc(var(--m3-shape-corner-md) - 2px)' }}
              />
            </div>
          </div>

          <div className="h-8 w-px mx-2 opacity-50" style={{ background: 'var(--m3-outline-variant)' }}></div>

          {/* Workspace ID Badge & Sync Status */}
          {!USE_LOCAL_STORAGE && (
            <div className="flex items-center gap-2 pl-2 pr-3 py-1.5 rounded-full shadow-sm group relative" style={{ background: 'var(--m3-primary-container)', color: 'var(--m3-on-primary-container)' }}>
              <div className="w-2.5 h-2.5 rounded-full animate-pulse" style={{ background: '#34d399' }} title="Cloud Synced"></div>
              <span className="text-xs font-bold uppercase tracking-wider">{appId}</span>

              <button
                onClick={() => {
                  const url = new URL(window.location);
                  url.searchParams.set('project', appId);
                  navigator.clipboard.writeText(url.toString());
                  showAlert("プロジェクト共有リンクをクリップボードにコピーしました");
                }}
                className="ml-1 p-1 rounded-full hover:bg-black/10 transition-colors"
                title="共有リンクをコピー"
              >
                <LinkIcon size={14} />
              </button>
            </div>
          )}

          <div className="flex p-1 rounded-full transition-all" style={{ border: '1px solid var(--m3-outline)', background: 'var(--m3-surface)' }}>
            <button
              onClick={() => { setViewMode('list'); setActiveSheetId(null); setIsPageSelectionMode(false); setIsLabelSelectionMode(false); }}
              onMouseEnter={(e) => showQuickHelp(e, '詳細', 'ページ単位で編集する表示に切り替えます。')}
              onMouseLeave={hideQuickHelp}
              className={`flex items-center gap-2 px-5 py-2 text-sm font-medium rounded-full transition-all duration-300 whitespace-nowrap`}
              style={viewMode === 'list' || viewMode === 'single' ? { background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
            >
              <List size={18} /> <span className="hidden sm:inline">詳細</span>
            </button>
            <button
              onClick={() => { setViewMode('overview'); setActiveSheetId(null); setIsPageSelectionMode(false); setIsLabelSelectionMode(false); }}
              onMouseEnter={(e) => showQuickHelp(e, '全体', '全ページを一覧で表示します。コマの全体把握に使います。')}
              onMouseLeave={hideQuickHelp}
              className={`flex items-center gap-2 px-5 py-2 text-sm font-medium rounded-full transition-all duration-300 whitespace-nowrap`}
              style={viewMode === 'overview' ? { background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
            >
              <Grid size={18} /> <span className="hidden sm:inline">全体</span>
            </button>
          </div>

          {/* Page Selection Mode Toggle */}
          <button
            onClick={togglePageSelectionMode}
            onMouseEnter={(e) => showQuickHelp(e, '選択モード', '複数ページを選択して、入れ替え・画像解除・削除を行います。')}
            onMouseLeave={hideQuickHelp}
            className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-full transition-all duration-300 border ml-3 whitespace-nowrap ${isPageSelectionMode ? 'bg-indigo-600 text-white border-indigo-600 shadow-md' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
            title="複数ページを選択して削除"
          >
            <CheckSquare size={14} strokeWidth={2.5} /> <span className="hidden sm:inline">選択モード</span>
          </button>

          <button
            onClick={() => setIsQuickHelpMode((prev) => !prev)}
            className={`ml-1 w-9 h-9 rounded-full border text-[14px] font-extrabold leading-none transition-all ${isQuickHelpMode
              ? 'bg-sky-600 text-white border-sky-500 ring-4 ring-sky-300/50 shadow-[0_0_20px_rgba(56,189,248,0.55)]'
              : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}
            title="クイックヘルプ"
          >
            Q
          </button>

          {isPageSelectionMode && (
            <div className="flex items-center gap-2 animate-in fade-in slide-in-from-left-4 duration-300 bg-white/50 backdrop-blur-sm px-2 py-1 rounded-xl border border-slate-200/50">
              <button
                onClick={handleSelectAllPages}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-bold whitespace-nowrap bg-white text-indigo-600 border border-indigo-100 hover:bg-indigo-50 hover:shadow-md`}
              >
                <CheckSquare size={14} />
                {selectedSheetIds.size === displaySheets.length && displaySheets.length > 0 ? '全解除' : '全選択'}
              </button>
              <span className="text-sm font-bold text-slate-600 ml-1 mr-2 whitespace-nowrap">
                {selectedSheetIds.size} / {displaySheets.length}
              </span>
              <button
                onClick={handleSwapPages}
                disabled={selectedSheetIds.size !== 2}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedSheetIds.size === 2 ? 'bg-indigo-500 text-white hover:bg-indigo-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
                title="選択した2つのページを入れ替え"
              >
                <ArrowLeftRight size={14} /> 入れ替え
              </button>
              <button
                onClick={handleBulkClearImages}
                disabled={selectedSheetIds.size === 0}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedSheetIds.size > 0 ? 'bg-amber-500 text-white hover:bg-amber-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
                title="選択したページの画像を全て外す"
              >
                <X size={14} /> 画像解除
              </button>
              <button
                onClick={handleBulkDelete}
                disabled={selectedSheetIds.size === 0}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all shadow-sm font-medium whitespace-nowrap ${selectedSheetIds.size > 0 ? 'bg-rose-500 text-white hover:bg-rose-600 hover:shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
              >
                <Trash2 size={14} /> 削除
              </button>
            </div>
          )}

          {!isPageSelectionMode && (viewMode === 'list' || viewMode === 'single') && (
            <div className={`flex items-center gap-1 ml-2 rounded-xl p-1 border ${isSalesMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-100/50 border-slate-200/50'}`}>
              <button
                onClick={toggleMergeMode}
                onMouseEnter={(e) => showQuickHelp(e, 'コマ結合', 'コマの結合/分離モードを切り替えます。複数コマを選択して結合できます。')}
                onMouseLeave={hideQuickHelp}
                className={`flex items-center gap-2 px-3 py-2 text-sm font-medium rounded-lg transition-all duration-200 whitespace-nowrap ${isMergeMode ? 'bg-indigo-100 text-indigo-700 shadow-inner' : (isSalesMode ? 'text-slate-300 hover:bg-slate-700' : 'text-slate-600 hover:bg-white hover:shadow-sm')}`}
                title="コマ結合・分離モードの切り替え"
              >
                <LinkIcon size={16} /> <span className="hidden lg:inline">コマ結合</span>
              </button>

              {isMergeMode && (
                <>
                  <div className="w-px h-6 bg-slate-300 mx-2 opacity-50"></div>
                  <button
                    onClick={handleMerge}
                    disabled={!canMerge}
                    className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all font-medium whitespace-nowrap ${canMerge ? 'bg-white shadow text-emerald-600 hover:text-emerald-700 hover:shadow-md' : 'text-slate-400 cursor-not-allowed'}`}
                    title="選択したコマを結合（長方形のみ）"
                  >
                    <Merge size={16} /> 結合
                  </button>
                  <button
                    onClick={handleSplit}
                    disabled={!canSplit}
                    className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded-lg transition-all font-medium whitespace-nowrap ${canSplit ? 'bg-white shadow text-amber-600 hover:text-amber-700 hover:shadow-md' : 'text-slate-400 cursor-not-allowed'}`}
                    title="選択したコマを分離"
                  >
                    <Split size={16} /> 分離
                  </button>
                </>
              )}
            </div>
          )}

          {/* 実績モード Toggle - 詳細表示時のみ */}
          {!isPageSelectionMode && (viewMode === 'list' || viewMode === 'single') && (
            <button
              onClick={() => setIsSalesMode(!isSalesMode)}
              onMouseEnter={(e) => showQuickHelp(e, '実績モード', '販売数量データをパネル上に重ねて表示します。再度押すと解除します。')}
              onMouseLeave={hideQuickHelp}
              className={`flex items-center gap-2 px-3 py-2 rounded-xl border text-sm font-bold transition-all duration-300 ml-2 whitespace-nowrap
                 ${isSalesMode
                  ? 'bg-emerald-500/10 border-emerald-500 text-emerald-400 shadow-[0_0_15px_rgba(16,185,129,0.3)]'
                  : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}`}
              title="実績モード切替（詳細表示時のみ）"
            >
              <BarChart2 size={20} />
              <span className="hidden xl:inline">実績モード {isSalesMode ? 'ON' : 'OFF'}</span>
            </button>
          )}
        </div>

        <div className="flex items-center gap-2 flex-shrink-0 ml-4">
          {/* Zoom Controls (Fixed bottom-left) - M3 Style */}
          {/* Zoom Controls */}
          {!isPageSelectionMode && (viewMode === 'list' || viewMode === 'single') && (
            <div className="flex items-center gap-1 bg-slate-100 rounded-lg p-1 mr-2 opacity-0 animate-in fade-in slide-in-from-right-4 duration-500" style={{ opacity: 1 }}>
              <button
                onClick={() => setZoomScale(s => Math.max(0.5, s - 0.1))}
                className="p-1.5 hover:bg-white hover:shadow-sm rounded-md text-slate-500 transition-all active:scale-95"
                title="縮小"
              >
                <ZoomOut size={16} />
              </button>
              <span className="text-xs font-mono font-bold w-12 text-center text-slate-600 select-none">
                {Math.round(zoomScale * 100)}%
              </span>
              <button
                onClick={() => setZoomScale(s => Math.min(1.5, s + 0.1))}
                className="p-1.5 hover:bg-white hover:shadow-sm rounded-md text-slate-500 transition-all active:scale-95"
                title="拡大"
              >
                <ZoomIn size={16} />
              </button>
            </div>
          )}

          {!isPageSelectionMode && viewMode === 'overview' && (
            <button
              onClick={() => setHighlightLabels(!highlightLabels)}
              onMouseEnter={(e) => showQuickHelp(e, 'ラベル強調', 'ラベルが1つ以上あるコマを緑色で強調表示します。もう一度押すと解除します。')}
              onMouseLeave={hideQuickHelp}
              className={`flex items-center gap-2 px-3 py-2 border rounded-xl text-sm font-medium transition-all duration-200 whitespace-nowrap ${highlightLabels ? 'bg-emerald-500 text-white border-emerald-600 shadow-md shadow-emerald-200' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              title="自由ラベルがあるコマを緑色で強調表示"
            >
              <Tag size={16} /> <span>ラベル強調</span>
            </button>
          )}

          <button
            onClick={() => setHighlightEmpty(!highlightEmpty)}
            onMouseEnter={(e) => showQuickHelp(e, '空き強調', '空きコマを赤色で強調表示します。全体表示時の確認に使います。')}
            onMouseLeave={hideQuickHelp}
            className={`flex items-center gap-2 px-3 py-2 border rounded-xl text-sm font-medium transition-all duration-200 whitespace-nowrap ${highlightEmpty ? 'bg-rose-500 text-white border-rose-600 shadow-md shadow-rose-200' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
            title="空きコマを赤色で強調表示"
          >
            <AlertCircle size={16} /> <span>空き強調</span>
          </button>

          <button
            onClick={handleExportCSV}
            onMouseEnter={(e) => showQuickHelp(e, '出力', '現在のページ情報をCSVで出力します。外部共有やバックアップに使えます。')}
            onMouseLeave={hideQuickHelp}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-xl hover:bg-emerald-100 text-sm font-medium transition-all duration-200 shadow-sm hover:shadow whitespace-nowrap"
            title="ページ情報をCSVでダウンロード"
          >
            <FileSpreadsheet size={16} /> <span>出力</span>
          </button>

          <div className="h-8 w-px bg-black/5 mx-2"></div>

          <div className="flex items-center gap-2 text-[10px] font-bold text-emerald-600 bg-emerald-500/10 px-3 py-1.5 rounded-full border border-emerald-500/10 whitespace-nowrap">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]"></div>
            Online
          </div>
        </div>
      </div>

      <div className="flex flex-1 overflow-hidden relative">
        <Sidebar
          isOpen={sidebarOpen}
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
          tempItems={tempItems}
          excludedItems={excludedItems}
          onMoveToTemp={handleMoveToTemp}
          onDeleteFromTemp={handleDeleteFromTemp}
          onMoveToExcluded={handleMoveToExcluded}
          onDeleteFromExcluded={handleDeleteFromExcluded}
          onExportExcludedCSV={handleExportExcludedCSV}
          onBulkDeleteExcluded={handleBulkDeleteExcluded}
          onMoveToStock={handleMoveToStock}
          imageDataById={imageDataById}
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
            <div className={`flex justify-between items-center gap-2 flex-wrap ${viewMode === 'single'
              ? `sticky top-0 z-40 rounded-xl border px-2 py-1.5 backdrop-blur ${isSalesMode
                ? 'bg-slate-900/85 border-slate-700 shadow-lg shadow-slate-900/20'
                : 'bg-white/90 border-slate-200 shadow-lg shadow-slate-200/70'}`
              : ''
              }`}>
              <div className="flex items-center gap-2 bg-white px-2.5 py-1.5 rounded-lg shadow-sm border border-slate-100">
                <span className="text-xs font-bold text-slate-600">ジャンル:</span>
                <select
                  value={genreFilter}
                  onChange={(e) => setGenreFilter(e.target.value)}
                  className="bg-slate-50 border-none rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 text-slate-700 font-medium cursor-pointer hover:bg-slate-100 transition-colors"
                >
                  <option value="all">全て表示</option>
                  {GENRES.map(g => (
                    <option key={g.id} value={g.id}>{g.label}</option>
                  ))}
                </select>
              </div>

              {/* Page Nav */}
              <div className={`flex items-center gap-2 px-2.5 py-1.5 rounded-lg shadow-sm border ${isSalesMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-100'}`}>
                <button
                  onClick={() => handleNavigatePage('prev')}
                  disabled={viewMode !== 'single' || currentIndex <= 0}
                  className="p-1.5 hover:bg-slate-100 rounded-md text-slate-500 disabled:opacity-30 disabled:cursor-not-allowed transition-all active:scale-95"
                >
                  <ChevronLeft size={18} />
                </button>
                <div className="flex flex-col items-center min-w-[5rem]">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider leading-none mb-0.5">Page</span>
                  <span className="font-mono text-base font-bold leading-none flex items-baseline">
                    {viewMode === 'single' && activeSheetId ? currentIndex + 1 : '-'}
                    <span className="text-slate-400 text-xs mx-1 font-normal">/</span>
                    <span className="text-sm text-slate-500 font-medium">{currentList.length}</span>
                  </span>
                </div>
                <button
                  onClick={() => handleNavigatePage('next')}
                  disabled={viewMode !== 'single' || currentIndex === -1 || currentIndex >= currentList.length - 1}
                  className="p-1.5 hover:bg-slate-100 rounded-md text-slate-500 disabled:opacity-30 disabled:cursor-not-allowed transition-all active:scale-95"
                >
                  <ChevronRight size={18} />
                </button>
              </div>

              {/* Header Controls inside Detail Area */}
              <div className="flex items-center gap-2 flex-wrap justify-end">
                {viewMode === 'single' && (
                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      onClick={handleBulkDeletePageLabels}
                      onMouseEnter={(e) => showQuickHelp(e, 'ラベル一括削除', 'このページの自由ラベルをまとめて削除します。確認後に実行されます。')}
                      onMouseLeave={hideQuickHelp}
                      disabled={activeSheetLabelCount === 0}
                      className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg shadow-sm transition-all active:scale-95 font-bold text-[11px] whitespace-nowrap ${activeSheetLabelCount > 0
                        ? 'bg-rose-600 text-white hover:bg-rose-700'
                        : 'bg-slate-200 text-slate-500 cursor-not-allowed'
                        }`}
                    >
                      <Trash2 size={14} strokeWidth={3} />
                      ラベル一括削除
                    </button>

                    <button
                      onClick={() => setIsLabelSelectionMode(!isLabelSelectionMode)}
                      onMouseEnter={(e) => showQuickHelp(e, 'ラベル追加', '自由ラベル配置モードをON/OFFします。ON中はコマ内クリックでラベルを配置できます。')}
                      onMouseLeave={hideQuickHelp}
                      className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg shadow-sm transition-all active:scale-95 font-bold text-[11px] whitespace-nowrap ${isLabelSelectionMode
                        ? 'bg-emerald-600 text-white shadow-emerald-200 ring-2 ring-emerald-500/30'
                        : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-50'
                        }`}
                    >
                      <Tag size={14} strokeWidth={3} className={isLabelSelectionMode ? 'animate-pulse' : ''} />
                      {isLabelSelectionMode ? '配置中...' : 'ラベル追加'}
                    </button>
                  </div>
                )}

                {/* Add Page Button */}
                <button
                  onClick={handleAddSheet}
                  onMouseEnter={(e) => showQuickHelp(e, '+ページ追加', '新しいページを末尾に追加します。')}
                  onMouseLeave={hideQuickHelp}
                  className="flex items-center gap-1.5 bg-indigo-600 text-white px-3 py-1.5 rounded-lg shadow-sm shadow-indigo-200 hover:bg-indigo-700 transition-all active:scale-95 font-bold text-[11px] whitespace-nowrap"
                >
                  <Plus size={14} strokeWidth={3} /> +ページ追加
                </button>
              </div>

            </div>

            <div
              className={`relative z-10 ${viewMode === 'overview' ? 'grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-8' : 'flex flex-col gap-12 items-center pb-32'}`}
              style={{
                transform: `scale(${zoomScale})`,
                transformOrigin: 'top center',
                minHeight: zoomScale > 1 ? `${zoomScale * 100}%` : 'auto'
              }}
            >
              {displaySheets.map((sheet) => {
                const isPageSelected = selectedSheetIds.has(sheet.id);
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

                    <div className="flex flex-col gap-3">
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

                      <div className={`relative ${isPageSelectionMode ? 'pointer-events-none' : ''}`}>
                        <Sheet
                          sheet={sheet}
                          index={sheets.findIndex(s => s.id === sheet.id)}
                          panels={sheet.panels}
                          updatePanel={handlePanelUpdateWithCheck}
                          isOverview={viewMode === 'overview'}
                          zoomScale={zoomScale}
                          selection={selection}
                          onSelectPanel={isMergeMode ? handleSelectPanel : undefined}
                          onDeleteSheet={handleDeleteSheet}
                          highlightEmpty={highlightEmpty}
                          highlightLabels={highlightLabels}
                          onMovePanel={handleMovePanel}
                          isSalesMode={isSalesMode}
                          salesData={salesData}
                          onHoverSales={handleHoverSales}
                          onLeaveSales={handleLeaveSales}
                          imageDataById={imageDataById}
                          isLabelMode={isLabelSelectionMode}
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

      {isQuickHelpMode && quickHelpPopup && (
        <div
          className="fixed z-[120] pointer-events-none"
          style={{ left: quickHelpPopup.x, top: quickHelpPopup.y, transform: 'translateX(-50%)' }}
        >
          <div className="min-w-[320px] max-w-[460px] rounded-2xl border border-sky-200 bg-white/95 backdrop-blur-sm px-4 py-3 shadow-xl">
            <p className="text-[13px] font-bold text-sky-700">{quickHelpPopup.title}</p>
            <p className="text-[12px] leading-relaxed text-slate-700 mt-1.5">{quickHelpPopup.description}</p>
          </div>
        </div>
      )}

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
      />

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
