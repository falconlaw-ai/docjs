import { expect, test, type Page } from "@playwright/test";
import { waitForCommittedMode } from "./helpers";

async function waitForPreview(page: Page) {
  await waitForCommittedMode(page, "original");
}

async function setZoomConfiguration(
  page: Page,
  values: { min?: string; max?: string; default?: string },
) {
  await page.getByRole("button", { name: "Viewer settings" }).click();
  if (values.min !== undefined) await page.getByLabel("Minimum zoom").fill(values.min);
  if (values.max !== undefined) await page.getByLabel("Maximum zoom").fill(values.max);
  if (values.default !== undefined) await page.getByLabel("Default zoom").fill(values.default);
  await page.keyboard.press("Escape");
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
  await expect(zoom).toHaveAttribute("min", "20");
  await expect(zoom).toHaveAttribute("max", "200");

  await page.goto(
    "/?fixture=consulting-markdown&minZoom=-50&maxZoom=500&defaultZoom=500",
  );
  await waitForPreview(page);
  await expect(zoom).toHaveValue("200");
  await expect(zoom).toHaveAttribute("min", "20");
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

  await page.getByRole("button", { name: "Viewer settings" }).click();
  await page.getByLabel("Minimum zoom").fill("100");
  await expect(zoom).toHaveValue("125");
  await page.getByLabel("Minimum zoom").fill("140");
  await expect(zoom).toHaveValue("140");
  await page.keyboard.press("Escape");
  await page.getByRole("button", { name: "Fit width" }).click();
  await page.getByRole("button", { name: "Viewer settings" }).click();
  await page.getByLabel("Maximum zoom").fill("100");
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
  await page.getByRole("button", { name: "Viewer settings" }).click();
  await page.getByLabel("Minimum zoom").fill("140");
  await expect(zoom).toHaveValue("140");
  await page.waitForTimeout(350);
  expect(await visibleTextAtTop(page)).toBe(before);
});

test("fits only oversized zoom when entering Review", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx&defaultZoom=150");
  await waitForPreview(page);

  const zoom = page.getByLabel("Zoom percentage");
  await expect(zoom).toHaveValue("150");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForCommittedMode(page, "review");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );

  await page.getByRole("button", { name: "Original", exact: true }).click();
  await waitForCommittedMode(page, "original");
  await zoom.fill("50");
  await zoom.press("Enter");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForCommittedMode(page, "review");
  await expect(zoom).toHaveValue("50");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "false",
  );
});

test("selects fit for physical overflow even when the minimum limits it", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto(
    "/?fixture=consulting-docx&minZoom=150&defaultZoom=150",
  );
  await waitForPreview(page);

  const zoom = page.getByLabel("Zoom percentage");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForCommittedMode(page, "review");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  await expect(zoom).toHaveValue("150");
});

test("keeps a manual Review zoom after the entry decision", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/?fixture=consulting-docx&defaultZoom=150");
  await waitForPreview(page);
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await waitForCommittedMode(page, "review");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );

  const zoom = page.getByLabel("Zoom percentage");
  await zoom.fill("180");
  await zoom.press("Enter");
  await page.getByTestId("item-add").click();
  await page.setViewportSize({ width: 1100, height: 800 });
  await expect(zoom).toHaveValue("180");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "false",
  );
});

for (const destination of ["Original", "Final"] as const) {
  test(`preserves Review's effective percentage when leaving fit width for ${destination}`, async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.goto("/?fixture=consulting-docx&defaultZoom=150");
    await waitForPreview(page);
    await page.getByRole("button", { name: "Review", exact: true }).click();
    await waitForCommittedMode(page, "review");

    const zoom = page.getByLabel("Zoom percentage");
    const reviewPercentage = await zoom.inputValue();
    await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
      "aria-pressed",
      "true",
    );
    await page.getByRole("button", { name: destination, exact: true }).click();
    await waitForCommittedMode(page, destination.toLowerCase() as "original" | "final");
    await expect(zoom).toHaveValue(reviewPercentage);
    await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
      "aria-pressed",
      "false",
    );
  });
}

test("waits for delayed Review preparation before applying its width", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto(
    "/?fixture=consulting-markdown&defaultZoom=150&delay=700",
  );
  await waitForPreview(page);

  const zoom = page.getByLabel("Zoom percentage");
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await expect(page.getByRole("button", { name: "Review", exact: true })).toHaveAttribute(
    "aria-busy",
    "true",
  );
  await expect(zoom).toHaveValue("150");
  await waitForCommittedMode(page, "review");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
});

test("lets an explicit default change replace a pending Review decision", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto(
    "/?fixture=consulting-markdown&defaultZoom=150&delay=1500",
  );
  await waitForPreview(page);
  await page.getByRole("button", { name: "Review", exact: true }).click();
  await expect(page.getByRole("button", { name: "Review", exact: true })).toHaveAttribute(
    "aria-busy",
    "true",
  );
  await setZoomConfiguration(page, { default: "60" });

  await waitForCommittedMode(page, "review");
  await expect(page.getByLabel("Zoom percentage")).toHaveValue("60");
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "false",
  );
});
