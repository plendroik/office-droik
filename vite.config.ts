import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
// @ts-expect-error -- düz JS sunucu modülü, tip tanımı yok
import { apiMiddleware, pdfjsAssetsMiddleware } from './server/api.js';

export default defineConfig({
  plugins: [
    react(),
    {
      name: 'office-droik-api',
      configureServer(server) {
        server.middlewares.use('/api', apiMiddleware());
        server.middlewares.use('/pdfjs', pdfjsAssetsMiddleware());
      },
    },
  ],
  server: { port: 5173, host: '127.0.0.1' },
  build: { chunkSizeWarningLimit: 8000 },
});
