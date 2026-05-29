import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  define: {
    // 빌드는 UTC CI에서 돌지만 표기는 한국시(KST, UTC+9)로 통일
    __BUILD_TIME__: JSON.stringify(
      new Date(Date.now()+32400000).toISOString().slice(0,16).replace('T',' ')+' KST'
    ),
  },
  build: {
    target: 'es2020',
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules/react-dom') || id.includes('node_modules/react/')) return 'vendor-react';
          if (id.includes('node_modules/recharts') || id.includes('node_modules/d3-') || id.includes('node_modules/victory-')) return 'vendor-recharts';
          if (id.includes('node_modules/@supabase')) return 'vendor-supabase';
          if (id.includes('node_modules/dayjs')) return 'vendor-dayjs';
        },
      },
    },
  },
})
