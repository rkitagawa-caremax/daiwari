import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    // 本番ビルドの最適化設定
    target: 'es2020',
    // esbuild（Viteデフォルト）を使用 - terserより高速
    minify: 'esbuild',
    rollupOptions: {
      output: {
        // コード分割：firebase/lucide を別チャンクに分離
        manualChunks: {
          'vendor-firebase': ['firebase/app', 'firebase/auth', 'firebase/firestore'],
          'vendor-ui': ['lucide-react']
        }
      }
    },
    // チャンクサイズ警告閾値を調整
    chunkSizeWarningLimit: 1000
  },
  // 本番ビルドでconsole.log等を削除
  esbuild: {
    drop: ['console', 'debugger']
  },
  // 開発サーバー最適化
  server: {
    hmr: {
      overlay: true
    }
  }
})
