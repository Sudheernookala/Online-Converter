import { defineConfig } from 'vite';

// Relative base so the build works at https://<user>.github.io/<repo>/
export default defineConfig({
  base: './',
  build: { chunkSizeWarningLimit: 3000 },
});
