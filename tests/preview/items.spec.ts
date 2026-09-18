import { expect, test } from "@playwright/test";
import { waitForCommittedMode } from "./helpers";

test("Final projects every supplied resolvable redline and hides comments", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await expect(page.getByTestId("source-digest")).not.toHaveText("pending");
  const originalDigest = await page.getByTestId("source-digest").textContent();
  await page.getByRole("button", { name: "Final" }).click();

  await waitForCommittedMode(page, "final");
  await expect(page.getByLabel("Review comments and redlines")).toHaveCount(0);
  await expect(page.getByTestId("supplied-redline-count")).toHaveText("2");
  const finalProjection = await page.locator(".docx-preview").evaluate((host) => {
    const root = host.shadowRoot;
    return {
      text: root?.textContent ?? "",
      decorations:
        root?.querySelectorAll(
          "ins, del, [data-docx-comment-ids], [data-docx-entity-id]",
        ).length ?? -1,
    };
  });
  expect(finalProjection.text).toContain("Consultant acts as an independent professional");
  expect(finalProjection.text).toContain("Company will pay the agreed fees");
  expect(finalProjection.text).not.toContain("Consultant is an independent contractor");
  expect(finalProjection.decorations).toBe(0);

  await page.getByRole("button", { name: "Original" }).click();
  await waitForCommittedMode(page, "original");
  expect(await page.getByTestId("source-digest").textContent()).toBe(originalDigest);
  expect(
    await page
      .locator(".docx-preview")
      .evaluate((host) => host.shadowRoot?.textContent ?? ""),
  ).toContain("Consultant is an independent contractor");
});

test("diagnostics stay visible instead of silently dropping invalid items", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();
  await waitForCommittedMode(page, "review");

  const invalid = page.locator('[data-review-item-id="invalid-anchor"]');
  await expect(invalid).toBeVisible({ timeout: 30_000 });
  await expect(invalid.locator("[data-review-diagnostic]")).toBeVisible();
});

test("item updates retain selection and item removal clears it", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();
  await waitForCommittedMode(page, "review");
  const selected = page.locator('[data-review-item-id="redline-consultant"]');
  await expect(selected).toBeVisible({ timeout: 30_000 });
  await selected.click();
  await expect(selected).toHaveAttribute("aria-pressed", "true");

  await page.getByTestId("item-update").click();
  await expect(page.getByTestId("item-version")).toHaveText("1");
  await expect(selected).toHaveAttribute("aria-pressed", "true");
  await expect
    .poll(() =>
      page
        .locator(".docx-preview")
        .evaluate((host) => host.shadowRoot?.textContent ?? ""),
    )
    .toContain(
      "Company will pay all agreed professional fees after receiving a valid invoice",
    );

  await page.getByTestId("item-update").click();
  await expect(page.getByTestId("item-version")).toHaveText("2");
  await expect(selected).toHaveAttribute("aria-pressed", "true");
  await expect
    .poll(() =>
      page
        .locator(".docx-preview")
        .evaluate((host) => host.shadowRoot?.textContent ?? ""),
    )
    .toContain("Company will pay the agreed fees");

  await page.getByTestId("item-add").click();
  await expect(page.locator('[data-review-item-id="comment-payment"]')).toBeVisible();
  await page.getByTestId("item-remove").click();
  await expect(selected).toHaveCount(0);
});
