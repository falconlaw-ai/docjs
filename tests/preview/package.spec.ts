import { expect, test } from "@playwright/test";

test("mounts the built React package and changes fixtures", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "Document preview example" })).toBeVisible();
  await expect(page.getByLabel("Example document")).toHaveValue("consulting-docx");
  await expect(page.getByLabel("Document preview")).toBeVisible();

  await page.getByLabel("Example document").selectOption("consulting-markdown");
  await expect(page.getByText("consulting_agreement.md", { exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "CONSULTING AGREEMENT" })).toBeVisible();
});

test("keeps zoom controls within the supported range", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown");
  const zoom = page.getByLabel("Zoom percentage");

  await zoom.fill("5");
  await zoom.press("Enter");
  await expect(zoom).toHaveValue("25");

  await zoom.fill("350");
  await zoom.press("Enter");
  await expect(zoom).toHaveValue("300");
  await page.getByRole("button", { name: "Fit width" }).click();
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute("aria-pressed", "true");
});
