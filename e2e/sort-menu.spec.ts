import { test, expect } from '@playwright/test';

test('header dropdown opens and sorts', async ({ page }) => {
  await page.goto('/');
  
  // wait for grid
  await expect(page.locator('.sheet')).toBeVisible();
  
  // click first column header dropdown trigger (▼)
  const trigger = page.locator('button.col-menu-trigger').first();
  await expect(trigger).toBeVisible();
  await trigger.click();
  
  // expect menu visible
  await expect(page.locator('.col-menu')).toBeVisible();
  
  // click Sort A→Z
  await page.locator('.col-menu-item', { hasText: 'Sort A→Z' }).click();
  
  // menu closes
  await expect(page.locator('.col-menu')).toBeHidden();
});
