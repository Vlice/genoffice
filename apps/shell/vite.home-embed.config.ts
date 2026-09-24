import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

const __dirname = fileURLToPath(new URL('.', import.meta.url))

/**
 * Standalone web build of GenOffice Home for MoreAI iframe embed.
 * Output: apps/shell/out/home-embed/
 */
export default defineConfig({
  root: resolve(__dirname, 'src/renderer'),
  base: './',
  plugins: [react()],
  build: {
    outDir: resolve(__dirname, 'out/home-embed'),
    emptyOutDir: true,
    rollupOptions: {
      input: resolve(__dirname, 'src/renderer/home-embed.html'),
    },
  },
  server: {
    port: Number(process.env.SHELL_HOME_EMBED_PORT) || 5200,
    strictPort: true,
  },
})
