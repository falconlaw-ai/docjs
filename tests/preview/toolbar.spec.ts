import { expect, test } from "@playwright/test";
import { waitForCommittedMode } from "./helpers";

test("shows only the current page count beside the format badge", async ({ page }) => {
  await page.goto("/?fixture=complex-docx");
  await waitForCommittedMode(page, "original");

  const documentInfo = page.locator(".docx-preview-document-info");
  const status = page.getByRole("status", { name: "Preview status" });
  await expect(status).toHaveText("17 pages");
  await expect(documentInfo).toContainText("msa_vantage_shield_complex.docx");
  await expect(documentInfo).not.toContainText("Original view");
  await expect(documentInfo).not.toContainText("Read only");
  expect(await status.evaluate((element) => element.previousElementSibling?.className)).toBe(
    "docx-preview-format-badge",
  );

  await page.getByRole("button", { name: "Review", exact: true }).click();
  await expect(status).toHaveCount(0);
  await expect(page.getByRole("status", { name: "Preview update" })).toContainText(
    "Preparing preview",
  );
  await waitForCommittedMode(page, "review");
  await expect(status).toHaveText("17 pages");
});

test("settings control the viewer without resetting review content", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForCommittedMode(page, "original");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForCommittedMode(page, "review");

  const settingsButton = page.getByRole("button", { name: "Viewer settings" });
  await settingsButton.click();
  await expect(settingsButton).toHaveAttribute("aria-expanded", "true");
  await expect(page.getByLabel("Minimum zoom")).toBeFocused();

  await page.getByLabel("Show view controls").uncheck();
  await page.getByLabel("Show zoom controls").uncheck();
  await page.getByLabel("Show review panel").uncheck();
  await expect(page.getByRole("group", { name: "Document view" })).toHaveCount(0);
  await expect(page.getByRole("group", { name: "Document zoom" })).toHaveCount(0);
  await expect(page.getByLabel("Review comments and redlines")).toHaveCount(0);
  await expect(
    page.locator('.docx-preview [data-docx-entity-id="redline-consultant"]').first(),
  ).toBeVisible();
  await expect(page.locator(".docx-preview-layout")).not.toHaveClass(/docx-preview-layout--review/);

  await page.keyboard.press("Escape");
  await expect(settingsButton).toHaveAttribute("aria-expanded", "false");
  await expect(settingsButton).toBeFocused();
  await settingsButton.click();
  await page.getByRole("heading", { name: "Document preview example" }).click();
  await expect(settingsButton).toHaveAttribute("aria-expanded", "false");
});

test("settings persist across dismissal and document changes", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown");
  await waitForCommittedMode(page, "original");

  const settingsButton = page.getByRole("button", { name: "Viewer settings" });
  await settingsButton.click();
  await page.getByLabel("Minimum zoom").fill("40");
  await page.getByLabel("Maximum zoom").fill("150");
  await page.getByLabel("Default zoom").fill("80");
  await page.keyboard.press("Escape");

  const zoom = page.getByLabel("Zoom percentage");
  await expect(zoom).toHaveValue("80");
  await expect(zoom).toHaveAttribute("min", "40");
  await expect(zoom).toHaveAttribute("max", "150");
  await page.getByLabel("Example document").selectOption("complex-markdown");
  await waitForCommittedMode(page, "original");
  await expect(zoom).toHaveValue("80");

  await settingsButton.click();
  await expect(page.getByLabel("Minimum zoom")).toHaveValue("40");
  await expect(page.getByLabel("Maximum zoom")).toHaveValue("150");
  await expect(page.getByLabel("Default zoom")).toHaveValue("80");
});

test("keeps the settings popover inside narrow viewports", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 700 });
  await page.goto("/?fixture=consulting-markdown");
  await waitForCommittedMode(page, "original");

  const settingsButton = page.getByRole("button", { name: "Viewer settings" });
  const picker = page.getByLabel("Example document");
  await settingsButton.click();
  const popover = page.getByRole("group", { name: "Preview settings" });
  const [buttonBox, pickerBox, popoverBox] = await Promise.all([
    settingsButton.boundingBox(),
    picker.boundingBox(),
    popover.boundingBox(),
  ]);
  expect(buttonBox).not.toBeNull();
  expect(pickerBox).not.toBeNull();
  expect(popoverBox).not.toBeNull();
  expect(buttonBox!.x + buttonBox!.width).toBeLessThanOrEqual(pickerBox!.x);
  expect(popoverBox!.x).toBeGreaterThanOrEqual(0);
  expect(popoverBox!.x + popoverBox!.width).toBeLessThanOrEqual(320);
  expect(popoverBox!.y + popoverBox!.height).toBeLessThanOrEqual(700);
});
