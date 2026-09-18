import { expect, test } from "@playwright/test";

test("keeps Original visible while Review is prepared", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown&delay=700");
  await expect(page.getByRole("heading", { name: "CONSULTING AGREEMENT" })).toBeVisible();

  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.getByRole("button", { name: "Review" })).toHaveAttribute("aria-busy", "true");
  await expect(page.getByRole("heading", { name: "CONSULTING AGREEMENT" })).toBeVisible();
  await expect(page.locator(".docx-preview-status")).toContainText("Preparing review");

  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  await expect(page.getByRole("button", { name: "Final" })).toBeEnabled();
});

test("mounts controlled Review with the Original fallback", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown&initialMode=review&delay=700");

  await expect(page.getByRole("button", { name: "Review" })).toHaveAttribute("aria-pressed", "true");
  await expect(page.locator(".docx-preview-status")).toContainText("Preparing review");
  await expect(page.getByRole("heading", { name: "CONSULTING AGREEMENT" })).toBeVisible();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
});

test("shows preparation failure with retry while preserving Original", async ({ page }) => {
  await page.goto("/?fixture=preparation-error&delay=50");
  await page.getByRole("button", { name: "Review" }).click();

  await expect(page.getByRole("alert")).toContainText("The example backend rejected this preparation");
  await expect(page.getByRole("button", { name: "Retry preparation" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "CONSULTING AGREEMENT" })).toBeVisible();
  await page.getByRole("button", { name: "Retry preparation" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  await expect(page.getByTestId("preparation-count")).toHaveText("2");
});

test("ignores stale preparation when the original changes", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown&delay=900");
  await page.getByRole("button", { name: "Review" }).click();
  await page.getByLabel("Example document").selectOption("complex-markdown");

  await expect(page.getByText("msa_vantage_shield_complex.md", { exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Original" })).toHaveAttribute("aria-pressed", "true");
  await page.waitForTimeout(1_000);
  await expect(page.getByText("msa_vantage_shield_complex.md", { exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Original" })).toHaveAttribute("aria-pressed", "true");
});

test("reuses a prepared working document for Final", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown&delay=20");
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  await expect(page.getByTestId("preparation-count")).toHaveText("1");

  await page.getByRole("button", { name: "Final" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Final view", { timeout: 30_000 });
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  await expect(page.getByTestId("preparation-count")).toHaveText("1");
});
