/* global __firebase_config, __app_id */
import { initializeApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';

// --- Firebase Configuration / Local Storage Mode ---
// このモジュールは import 時に一度だけ評価される (App.jsx 先頭にあった初期化処理をそのまま移設)。

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
} catch {
  activeConfig = firebaseConfig;
}

// URL の ?mode=local|cloud を優先し、なければ前回の選択 (localStorage) を使う。既定は cloud。
export const resolveStorageMode = () => {
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

let useLocalStorage = resolveStorageMode() === 'local';

let app, auth, db;
export const DEFAULT_APP_ID = typeof __app_id !== 'undefined' ? __app_id : 'default-workspace';

if (!useLocalStorage) {
  try {
    app = initializeApp(activeConfig);
    auth = getAuth(app);
    db = getFirestore(app);
  } catch (e) {
    console.error('Firebase initialization failed, falling back to localStorage:', e);
    useLocalStorage = true; // エラー時はローカルモードに強制移行
  }
}

// 初期化結果を反映した最終的なモード。以降は変化しない。
export const USE_LOCAL_STORAGE = useLocalStorage;
export { app as firebaseApp, auth, db };

export const CLOUD_IMAGES_CACHE_KEY = 'cloudImagesCache';
export const CLOUD_SALES_CACHE_KEY = 'cloudSalesDataCache';
export const CLOUD_CACHE_TTL_MS = 6 * 60 * 60 * 1000;
export const LOCAL_WORK_LOGS_KEY = 'daiwari_work_activity_logs_v1';
