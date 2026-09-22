import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

// 빌드 결과를 index.html 한 파일로 묶어 오프라인에서도 더블클릭으로 실행할 수 있게 한다.
export default defineConfig({
  base: './',
  plugins: [react(), viteSingleFile()],
  server: {
    host: '0.0.0.0',
    port: 3000,
    allowedHosts: true,
  },
  build: { chunkSizeWarningLimit: 4000 },
});
