import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { nativeReportPlugin } from './scripts/native-report-plugin';

export default defineConfig(({ mode }) => ({
  plugins: [react(), nativeReportPlugin()],
  server: {
    watch: { ignored: ['**/output/**', '**/.local/**'] },
    port: mode === 'office' ? 3443 : 5173,
    strictPort: true,
    ...(mode === 'office'
      ? {
          https: {
            key: readFileSync(join(homedir(), '.office-addin-dev-certs', 'localhost.key')),
            cert: readFileSync(join(homedir(), '.office-addin-dev-certs', 'localhost.crt')),
          },
        }
      : {}),
  },
  preview: { port: 4173, strictPort: true },
  build: { target: 'es2022' },
}));
