import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules/react-dom') || id.includes('node_modules/react/')) return 'vendor-react';
          if (id.includes('node_modules/recharts') || id.includes('node_modules/d3-') || id.includes('node_modules/victory-')) return 'vendor-recharts';
          if (id.includes('node_modules/@supabase')) return 'vendor-supabase';
          // xlsx and papaparse are dynamic imports — Vite auto-splits them
        },
      },
    },
  },
})
