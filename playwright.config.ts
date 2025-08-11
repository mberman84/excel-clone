import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: 'e2e',
  timeout: 30_000,
  expect: { timeout: 5000 },
  use: { baseURL: 'http://localhost:4173', trace: 'on-first-retry' },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: { 
    // Build first so Playwright always serves the latest assets
    command: 'npm run build && npm run preview', 
    port: 4173, 
    reuseExistingServer: true 
  }
});
