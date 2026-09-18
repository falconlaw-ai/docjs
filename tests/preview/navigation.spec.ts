import { expect, test } from "@playwright/test";

test("renders default cards inside viewer-owned navigation wrappers", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });

  const card = page.locator('[data-review-item-id="comment-services"]');
  await expect(card).toBeVisible({ timeout: 30_000 });
  await card.click();
  await expect(card).toHaveAttribute("aria-pressed", "true");
  await card.press("Enter");
  await expect(card).toHaveAttribute("aria-pressed", "true");
});

test("card activation scrolls the document viewport to its anchor", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  const card = page.locator('[data-review-item-id="redline-fees"]');
  await expect(card).toBeVisible({ timeout: 30_000 });
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.evaluate((element) => {
    element.scrollTop = 0;
  });
  await card.click();
  await expect.poll(() => viewport.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
});

test("document activation selects and reveals the matching card", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });
  const card = page.locator('[data-review-item-id="redline-fees"]');
  const target = page
    .locator('.docx-preview [data-docx-entity-id="redline-fees"]')
    .first();

  await target.click();
  await expect(card).toHaveAttribute("aria-pressed", "true");
  await expect(card).toBeInViewport();
});

test("custom cards keep selection and expose consumer actions", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx&cards=custom");
  await page.getByRole("button", { name: "Review" }).click();
  await expect(page.locator(".docx-preview-status")).toContainText("Review view", { timeout: 30_000 });

  const card = page.locator('[data-review-item-id="redline-consultant"]');
  await expect(card).toBeVisible({ timeout: 30_000 });
  await expect(card.getByText("Consumer review card")).toBeVisible();
  await card.getByRole("button", { name: "Record consumer action" }).click();
  await expect(page.getByTestId("consumer-action")).toHaveText("redline-consultant");
  await expect(card).not.toHaveAttribute("aria-pressed", "true");

  await card.click();
  await expect(card).toHaveAttribute("aria-pressed", "true");
});

for (const key of ["Enter", "Space"] as const) {
  test(`custom-card actions handle ${key} without activating their wrapper`, async ({
    page,
  }) => {
    await page.goto("/?fixture=consulting-docx&cards=custom");
    await page.getByRole("button", { name: "Review" }).click();
    await expect(page.locator(".docx-preview-status")).toContainText("Review view", {
      timeout: 30_000,
    });

    const card = page.locator('[data-review-item-id="redline-consultant"]');
    const action = card.getByRole("button", { name: "Record consumer action" });
    await action.focus();
    await action.press(key);

    await expect(page.getByTestId("consumer-action")).toHaveText("redline-consultant");
    await expect(card).not.toHaveAttribute("aria-pressed", "true");
  });
}
