/**
 * IndexedDB Helper for Daiwari App
 * Solves the 5MB localStorage limit.
 */

const DB_NAME = 'daiwari_db';
const STORE_NAME = 'data';
const DB_VERSION = 1;

export const idbHelper = {
    db: null,

    async init() {
        if (this.db) return this.db;
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME);
                }
            };
            request.onsuccess = (e) => {
                this.db = e.target.result;
                resolve(this.db);
            };
            request.onerror = (e) => reject(e);
        });
    },

    async getItem(key) {
        const db = await this.init();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.get(key);
            request.onsuccess = () => resolve(request.result);
            request.onerror = (e) => reject(e);
        });
    },

    async setItem(key, value) {
        const db = await this.init();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readwrite');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.put(value, key);
            request.onsuccess = () => resolve();
            request.onerror = (e) => reject(e);
        });
    },

    generateId() {
        return `local_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }
};
