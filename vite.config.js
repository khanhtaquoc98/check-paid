import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 4321,
    open: true,
    proxy: {
      '/api': {
        target: 'https://dat-com-ivory.vercel.app',
        changeOrigin: true,
        secure: true,
      },
    },
  },
})
