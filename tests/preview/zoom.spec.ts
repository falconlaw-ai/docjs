import { expect, test, type Page } from "@playwright/test";

async function waitForPreview(page: Page) {
  await expect(page.locator(".docx-preview-status")).toContainText("Original view", {
    timeout: 30_000,
  });
}

async function setZoomConfiguration(
  page: Page,
  values: { min?: string; max?: string; default?: string },
) {
  if (values.min !== undefined) await page.getByLabel("Min %").fill(values.min);
  if (values.max !== undefined) await page.getByLabel("Max %").fill(values.max);
  if (values.default !== undefined) await page.getByLabel("Default %").fill(values.default);
}

async function visibleTextAtTop(page: Page) {
  return page.locator(".docx-preview-viewport").evaluate((viewport) => {
    const root = viewport.querySelector<HTMLElement>(".docx-preview")?.shadowRoot;
    const top = viewport.getBoundingClientRect().top;
    const normalize = (element: Element) =>
      (element.textContent ?? "").replace(/[\s\u200b-\u200d\ufeff\u00ad]+/g, "");
    return [...(root?.querySelectorAll<HTMLElement>("section.docx > article p") ?? [])]
      .map((element) => ({ element, bounds: element.getBoundingClientRect() }))
      .filter(
        ({ element, bounds }) =>
          bounds.bottom > top + 10 &&
          bounds.top < top + viewport.clientHeight &&
          normalize(element).length >= 24,
      )
      .sort((left, right) => left.bounds.top - right.bounds.top)[0]
      ?.element.textContent?.replace(/[\s\u200b-\u200d\ufeff\u00ad]+/g, "")
      .slice(0, 100) ?? "";
  });
}

test("uses fit width for absent, null, and nonfinite defaults", async ({ page }) => {
  for (const query of ["", "&defaultZoom=null", "&defaultZoom=Infinity"]) {
    await page.goto(`/?fixture=consulting-markdown${query}`);
    await waitForPreview(page);
    await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
      "aria-pressed",
      "true",
    );
  }
});

test("normalizes configured bounds and clamps the default zoom", async ({ page }) => {
  await page.goto(
    "/?fixture=consulting-markdown&minZoom=80&maxZoom=120&defaultZoom=500",
  );
  await waitForPreview(page);

  const zoom = page.getByLabel("Zoom percentage");
  await expect(zoom).toHaveValue("120");
  await expect(zoom).toHaveAttribute("min", "80");
  await expect(zoom).toHaveAttribute("max", "120");
  await expect(page.getByRole("button", { name: "Zoom in" })).toBeDisabled();
  await page.getByRole("button", { name: "Zoom out" }).click();
  await expect(zoom).toHaveValue("110");
  await zoom.fill("70");
  await zoom.press("Enter");
  await expect(zoom).toHaveValue("80");
  await expect(page.getByRole("button", { name: "Zoom out" })).toBeDisabled();

  await page.goto(
    "/?fixture=consulting-markdown&minZoom=190&maxZoom=50&defaultZoom=100",
  );
  await waitForPreview(page);
  await expect(zoom).toHaveValue("190");
  await expect(zoom).toHaveAttribute("min", "190");
  await expect(zoom).toHaveAttribute("max", "190");

  await page.goto(
    "/?fixture=consulting-markdown&minZoom=NaN&maxZoom=Infinity&defaultZoom=2000",
  );
  await waitForPreview(page);
  await expect(zoom).toHaveValue("200");
  await expect(zoom).toHaveAttribute("min", "25");
  await expect(zoom).toHaveAttribute("max", "200");

  await page.goto(
    "/?fixture=consulting-markdown&minZoom=-50&maxZoom=500&defaultZoom=500",
  );
  await waitForPreview(page);
  await expect(zoom).toHaveValue("200");
  await expect(zoom).toHaveAttribute("min", "25");
  await expect(zoom).toHaveAttribute("max", "200");
});

test("applies live configuration without remounting the document", async ({ page }) => {
  await page.addInitScript(() => {
    const original = File.prototype.text;
    Object.defineProperty(window, "__fileTextReads", {
      configurable: true,
      value: 0,
      writable: true,
    });
    File.prototype.text = function text() {
      const target = window as typeof window & { __fileTextReads: number };
      target.__fileTextReads += 1;
      return original.call(this);
    };
  });
  await page.goto("/?fixture=consulting-markdown");
  await waitForPreview(page);
  const zoom = page.getByLabel("Zoom percentage");

  await setZoomConfiguration(page, { min: "25", max: "150", default: "50" });
  await expect(zoom).toHaveValue("50");
  await zoom.fill("125");
  await zoom.press("Enter");
  await page.getByTestId("item-add").click();
  await expect(zoom).toHaveValue("125");

  await page.getByLabel("Min %").fill("100");
  await expect(zoom).toHaveValue("125");
  await page.getByLabel("Min %").fill("140");
  await expect(zoom).toHaveValue("140");
  await page.getByRole("button", { name: "Fit width" }).click();
  await page.getByLabel("Max %").fill("100");
  await expect(zoom).toHaveValue("140");
  await expect(zoom).toHaveAttribute("min", "140");
  await expect(zoom).toHaveAttribute("max", "140");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as typeof window & { __fileTextReads: number }).__fileTextReads,
      ),
    )
    .toBe(1);
});

test("reapplies default zoom for a new document identity", async ({ page }) => {
  await page.goto("/?fixture=consulting-markdown&defaultZoom=90");
  await waitForPreview(page);
  const zoom = page.getByLabel("Zoom percentage");
  await expect(zoom).toHaveValue("90");
  await zoom.fill("125");
  await zoom.press("Enter");

  await page.getByLabel("Open a document").setInputFiles(
    "tests/preview/fixtures/msa_vantage_shield_complex.md",
  );
  await expect(page.locator(".docx-preview-filename")).toHaveText(
    "msa_vantage_shield_complex.md",
  );
  await waitForPreview(page);
  await expect(zoom).toHaveValue("90");
});

test("preserves reading position when live bounds clamp zoom", async ({ page }) => {
  await page.goto("/?fixture=consulting-docx");
  await waitForPreview(page);
  const viewport = page.locator(".docx-preview-viewport");
  await viewport.evaluate((element) => {
    element.scrollTop = 0.42 * (element.scrollHeight - element.clientHeight);
  });
  await page.waitForTimeout(100);
  const before = await visibleTextAtTop(page);
  expect(before).not.toBe("");

  const zoom = page.getByLabel("Zoom percentage");
  await zoom.fill("100");
  await zoom.press("Enter");
  await page.getByLabel("Min %").fill("140");
  await expect(zoom).toHaveValue("140");
  await page.waitForTimeout(350);
  expect(await visibleTextAtTop(page)).toBe(before);
});

test("places the zoom configuration before the picker and wraps it on narrow screens", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto("/?fixture=consulting-markdown");
  const panel = page.getByRole("group", { name: "Zoom configuration" });
  const picker = page.getByLabel("Example document");
  await expect(panel).toBeVisible();
  const desktopPanel = await panel.boundingBox();
  const desktopPicker = await picker.boundingBox();
  expect(desktopPanel).not.toBeNull();
  expect(desktopPicker).not.toBeNull();
  expect(desktopPanel!.x + desktopPanel!.width).toBeLessThanOrEqual(desktopPicker!.x);

  await page.setViewportSize({ width: 600, height: 900 });
  const narrowPanel = await panel.boundingBox();
  const narrowPicker = await picker.boundingBox();
  expect(narrowPanel).not.toBeNull();
  expect(narrowPicker).not.toBeNull();
  const overlaps =
    narrowPanel!.x < narrowPicker!.x + narrowPicker!.width &&
    narrowPanel!.x + narrowPanel!.width > narrowPicker!.x &&
    narrowPanel!.y < narrowPicker!.y + narrowPicker!.height &&
    narrowPanel!.y + narrowPanel!.height > narrowPicker!.y;
  expect(overlaps).toBe(false);
  expect(narrowPanel!.x + narrowPanel!.width).toBeLessThanOrEqual(600);
  expect(narrowPicker!.x + narrowPicker!.width).toBeLessThanOrEqual(600);
});
