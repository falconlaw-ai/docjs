import { expect, test, type Page } from "@playwright/test";
import { waitForCommittedMode } from "./helpers";

type ReadingState = {
  progress: number;
  text: string;
};

async function waitForMode(page: Page, mode: "Original" | "Review" | "Final") {
  await waitForCommittedMode(page, mode.toLowerCase() as "original" | "review" | "final");
  await page.waitForTimeout(350);
}

async function setProgress(page: Page, progress: number): Promise<ReadingState> {
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.evaluate((element, nextProgress) => {
    element.scrollTop =
      nextProgress * Math.max(0, element.scrollHeight - element.clientHeight);
  }, progress);
  await page.waitForTimeout(50);
  return readingState(page);
}

async function readingState(page: Page): Promise<ReadingState> {
  return page.locator(".docx-preview-viewport").evaluate((viewport) => {
    const root = viewport.querySelector<HTMLElement>(".docx-preview")?.shadowRoot;
    const top = viewport.getBoundingClientRect().top;
    const normalize = (element: Element) =>
      (element.textContent ?? "")
        .replace(/[\s\u200b-\u200d\ufeff\u00ad]+/g, "")
        .replace(/^(?:\d+(?:\.\d+)*[.)]?|\([a-z]\)|[•▪])/, "");
    const block = [...(root?.querySelectorAll<HTMLElement>("section.docx > article p") ?? [])]
      .map((element) => ({ element, bounds: element.getBoundingClientRect() }))
      .filter(
        ({ element, bounds }) =>
          bounds.bottom > top + 10 &&
          bounds.top < top + viewport.clientHeight &&
          normalize(element).length >= 24,
      )
      .sort((left, right) => left.bounds.top - right.bounds.top)[0]?.element;
    return {
      progress:
        viewport.scrollTop /
        Math.max(1, viewport.scrollHeight - viewport.clientHeight),
      text: block ? normalize(block).slice(0, 100) : "",
    };
  });
}

test("restores each mode's saved reading position", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");

  const original = await setProgress(page, 0.28);
  expect(original.text).not.toBe("");
  await page.getByRole("button", { name: "Review" }).click();
  await waitForMode(page, "Review");

  const review = await setProgress(page, 0.57);
  expect(review.text).not.toBe("");
  await page.getByRole("button", { name: "Final" }).click();
  await waitForMode(page, "Final");

  const final = await setProgress(page, 0.18);
  expect(final.text).not.toBe("");
  await page.getByRole("button", { name: "Original" }).click();
  await waitForMode(page, "Original");
  expect((await readingState(page)).text).toBe(original.text);

  await page.getByRole("button", { name: "Review" }).click();
  await waitForMode(page, "Review");
  expect((await readingState(page)).text).toBe(review.text);

  await page.getByRole("button", { name: "Final" }).click();
  await waitForMode(page, "Final");
  expect((await readingState(page)).text).toBe(final.text);
});

test("keeps text continuity between fit and manual zoom", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");

  const beforeManual = await setProgress(page, 0.36);
  expect(beforeManual.text).not.toBe("");
  const zoom = page.getByLabel("Zoom percentage");
  await zoom.fill("75");
  await zoom.press("Enter");
  await expect
    .poll(() =>
      page.locator(".docx-scale-layer").evaluate((element) => element.style.transform),
    )
    .toBe("scale(0.75)");
  await page.waitForTimeout(350);
  expect((await readingState(page)).text).toBe(beforeManual.text);

  const beforeFit = await setProgress(page, 0.61);
  await page.getByRole("button", { name: "Fit width" }).click();
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  await page.waitForTimeout(350);
  expect((await readingState(page)).text).toBe(beforeFit.text);
});

test("starts a new document identity at the top", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  await setProgress(page, 0.72);

  await page.getByLabel("Open a document").setInputFiles(
    "tests/preview/fixtures/msa_vantage_shield_complex.docx",
  );
  await expect(page.locator(".docx-preview-filename")).toHaveText(
    "msa_vantage_shield_complex.docx",
  );
  await expect(page.locator(".docx-preview")).toHaveAttribute("aria-busy", "false", {
    timeout: 60_000,
  });
  await page.waitForTimeout(350);

  expect(
    await page
      .locator(".docx-preview-viewport")
      .evaluate((viewport) => viewport.scrollTop),
  ).toBeLessThan(5);
});

test("does not override direct scrolling during a mode change", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  await setProgress(page, 0.27);

  await page.getByRole("button", { name: "Review" }).click();
  await waitForMode(page, "Review");
  await setProgress(page, 0.64);

  await page.getByRole("button", { name: "Original" }).click();
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.hover();
  await page.mouse.wheel(0, -10_000);
  await waitForMode(page, "Original");

  expect((await readingState(page)).progress).toBeLessThan(0.05);
});
