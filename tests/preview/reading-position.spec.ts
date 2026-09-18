import { expect, test, type Page } from "@playwright/test";
import { waitForCommittedMode } from "./helpers";

type ReadingState = {
  progress: number;
  scrollTop: number;
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
      scrollTop: viewport.scrollTop,
      text: block ? normalize(block).slice(0, 100) : "",
    };
  });
}

async function openSettings(page: Page) {
  const button = page.getByRole("button", { name: "Viewer settings" });
  if ((await button.getAttribute("aria-expanded")) !== "true") await button.click();
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

test("preserves the fit-width reading anchor when the review panel is hidden and shown", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  const sourceDigestOutput = page.getByTestId("source-digest");
  await expect(sourceDigestOutput).not.toHaveText("pending");
  const sourceDigest = await sourceDigestOutput.textContent();
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForMode(page, "Review");

  const selected = page.locator('[data-review-item-id="redline-consultant"]');
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.evaluate((element) => {
    element.dataset.selectionScroll = "pending";
    element.addEventListener(
      "scrollend",
      () => {
        element.dataset.selectionScroll = "complete";
      },
      { once: true },
    );
  });
  await selected.click();
  await expect(selected).toHaveAttribute("aria-pressed", "true");
  await expect(viewport).toHaveAttribute("data-selection-scroll", "complete");
  const preview = page.locator(".docx-preview");
  await preview.evaluate((node) => {
    (window as typeof window & { __reviewPreview?: Element }).__reviewPreview = node;
  });
  const highlightCount = await page
    .locator(".docx-preview [data-docx-entity-id]")
    .count();
  expect(highlightCount).toBeGreaterThan(0);

  const before = await setProgress(page, 0.45);
  expect(before.text).not.toBe("");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  await openSettings(page);
  await page.getByLabel("Show review panel").uncheck();

  await expect(page.getByLabel("Review comments and redlines")).toHaveCount(0);
  await expect(page.locator(".docx-preview-shell")).toHaveAttribute(
    "data-committed-mode",
    "review",
  );
  await expect
    .poll(async () => (await readingState(page)).text)
    .toBe(before.text);
  expect(await page.locator(".docx-preview [data-docx-entity-id]").count()).toBe(
    highlightCount,
  );
  expect(
    await preview.evaluate(
      (node) =>
        node === (window as typeof window & { __reviewPreview?: Element }).__reviewPreview,
    ),
  ).toBe(true);

  await page.getByLabel("Show review panel").check();
  await expect(page.getByLabel("Review comments and redlines")).toBeVisible();
  await expect
    .poll(async () => (await readingState(page)).text)
    .toBe(before.text);
  await expect(selected).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("source-digest")).toHaveText(sourceDigest ?? "");
  expect(
    await preview.evaluate(
      (node) =>
        node === (window as typeof window & { __reviewPreview?: Element }).__reviewPreview,
    ),
  ).toBe(true);
});

test("leaves manual zoom stable across review panel changes", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForMode(page, "Review");
  const zoom = page.getByLabel("Zoom percentage");
  await zoom.fill("75");
  await zoom.press("Enter");
  await expect
    .poll(() =>
      page.locator(".docx-scale-layer").evaluate((element) => element.style.transform),
    )
    .toBe("scale(0.75)");
  await page.waitForTimeout(350);
  const before = await setProgress(page, 0.45);

  await openSettings(page);
  await page.getByLabel("Show review panel").uncheck();
  await expect(page.locator(".docx-scale-layer")).toHaveCSS(
    "transform",
    "matrix(0.75, 0, 0, 0.75, 0, 0)",
  );
  await expect.poll(async () => (await readingState(page)).scrollTop).toBe(before.scrollTop);
  await page.getByLabel("Show review panel").check();
  await expect.poll(async () => (await readingState(page)).scrollTop).toBe(before.scrollTop);
});

test("does not override direct scrolling during a review panel change", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForMode(page, "Review");
  await setProgress(page, 0.64);

  await openSettings(page);
  await page.getByLabel("Show review panel").uncheck();
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.hover();
  await page.mouse.wheel(0, -10_000);
  await page.waitForTimeout(350);

  expect((await readingState(page)).progress).toBeLessThan(0.05);
});

test("does not replace a pending mode restoration during a review panel change", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const arrayBuffer = File.prototype.arrayBuffer;
    File.prototype.arrayBuffer = function delayedArrayBuffer() {
      const result = arrayBuffer.call(this);
      const target = window as typeof window & { __delayNextDocxRead?: boolean };
      if (!target.__delayNextDocxRead) return result;
      target.__delayNextDocxRead = false;
      return new Promise((resolve, reject) => {
        window.setTimeout(() => result.then(resolve, reject), 700);
      });
    };
  });
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx");
  await waitForMode(page, "Original");
  const before = await setProgress(page, 0.45);
  expect(before.text).not.toBe("");
  await page.evaluate(() => {
    (window as typeof window & { __delayNextDocxRead?: boolean }).__delayNextDocxRead = true;
  });

  await page.getByRole("button", { name: "Review", exact: true }).click();
  await expect(page.locator(".docx-preview-render-slot.is-staged")).toHaveCount(1);
  await openSettings(page);
  await page.getByLabel("Show review panel").uncheck();
  await waitForMode(page, "Review");

  await expect
    .poll(async () => (await readingState(page)).text)
    .toBe(before.text);
});

test("preserves a ready Original fallback while Review preparation is pending", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-markdown&delay=3000");
  await waitForMode(page, "Original");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await expect(page.getByRole("button", { name: "Review", exact: true })).toHaveAttribute(
    "aria-busy",
    "true",
  );
  const before = await setProgress(page, 0.45);
  expect(before.progress).toBeGreaterThan(0.4);

  await openSettings(page);
  await page.getByLabel("Show review panel").uncheck();

  await expect(page.locator(".docx-preview-shell")).toHaveAttribute(
    "data-preview-mode",
    "review",
  );
  await expect(page.locator(".docx-preview-shell")).toHaveAttribute(
    "data-committed-mode",
    "original",
  );
  await expect
    .poll(async () => (await readingState(page)).progress)
    .toBeCloseTo(before.progress, 2);
});
