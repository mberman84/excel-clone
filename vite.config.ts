import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: { port: 5173 },
  // Vitest configuration
  test: {
    environment: 'jsdom',
    setupFiles: 'src/test/setup.ts',
    css: true,
    globals: true,
    coverage: {
      reporter: ['text', 'html'],
      reportsDirectory: 'coverage',
      exclude: ['node_modules', 'dist']
    }
  }
})
