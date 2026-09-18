import { expect, type Page } from "@playwright/test";

export async function waitForCommittedMode(
  page: Page,
  mode: "original" | "review" | "final",
) {
  await expect(
    page.locator(`.docx-preview-shell[data-committed-mode="${mode}"]`),
  ).toBeVisible({ timeout: 30_000 });
}
