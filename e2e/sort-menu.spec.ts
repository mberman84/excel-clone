import { test, expect } from '@playwright/test';

test('header dropdown opens and sorts', async ({ page }) => {
  await page.goto('/');
  
  // wait for grid
  await expect(page.locator('.header-row')).toBeVisible();
  
  // click first column header dropdown trigger (▼)
  const trigger = page.locator('button.header-col-menu-trigger').first();
  await trigger.click();
  
  // expect menu visible
  await expect(page.locator('.col-menu')).toBeVisible();
  
  // click Sort A→Z
  await page.getByRole('menuitem', { name: 'Sort A→Z' }).click();
  
  // menu closes
  await expect(page.locator('.col-menu')).toBeHidden();
});
