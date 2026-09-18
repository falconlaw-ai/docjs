import { createHash } from "node:crypto";

import { expect, test } from "@playwright/test";

function normalizedHash(parts: readonly string[]): {
  characters: number;
  sha256: string;
} {
  const text = parts.join(" ").replace(/\s+/gu, " ").trim();
  return {
    characters: text.length,
    sha256: createHash("sha256").update(text).digest("hex"),
  };
}

test("preserves the consulting baseline and its pagination corrections", async ({
  page,
}) => {
  await page.goto("/?fixture=consulting-docx");
  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 30_000,
  });

  const pages = page.locator(".docx-render-target section.docx");
  await expect(pages.first()).toBeVisible();
  const geometry = await pages.evaluateAll((nodes) =>
    nodes.map((node) => {
      const style = getComputedStyle(node);
      return {
        borderRadius: style.borderRadius,
        height: style.height,
        padding: style.padding,
        width: style.width,
      };
    }),
  );
  expect(geometry.length).toBeGreaterThan(0);
  expect(geometry).toEqual(
    Array.from({ length: geometry.length }, () => ({
      borderRadius: "3px",
      height: "1056px",
      padding: "86.4px",
      width: "816px",
    })),
  );

  const body = normalizedHash(
    await page
      .locator(".docx-render-target section.docx > article")
      .allTextContents(),
  );
  expect(body).toEqual({
    characters: 12_799,
    sha256: "11a23fc0e8012b67d3a52c8cfeb0aebf13ef262b8149917ac5ce553581e33c78",
  });

  const carriers = page.locator(".docx-section-carrier");
  await expect(carriers).not.toHaveCount(0);
  for (const carrier of await carriers.all()) {
    await expect
      .poll(() =>
        carrier.evaluate((node) => {
          const style = getComputedStyle(node);
          return [
            style.height,
            style.lineHeight,
            style.marginBottom,
            style.marginTop,
            style.minHeight,
          ];
        }),
      )
      .toEqual(["0px", "0px", "0px", "0px", "0px"]);
  }
});

test("preserves mixed A4 orientations in the complex DOCX", async ({ page }) => {
  await page.goto("/?fixture=complex-docx");
  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 60_000,
  });

  const pages = page.locator(".docx-render-target section.docx");
  await expect(pages.first()).toBeVisible();
  const sizes = await pages.evaluateAll((nodes) =>
    nodes.map((node) => {
      const style = getComputedStyle(node);
      return {
        height: Number.parseFloat(style.height),
        width: Number.parseFloat(style.width),
      };
    }),
  );
  expect(sizes.some(({ height, width }) => width > height)).toBe(true);
  expect(sizes.some(({ height, width }) => width < height)).toBe(true);
  for (const size of sizes) {
    expect(
      size.height === 1122.53 && size.width === 793.719 ||
        size.height === 793.719 && size.width === 1122.53,
    ).toBe(true);
  }
  await expect(
    page.getByText("MASTER SUBSCRIPTION AGREEMENT", { exact: true }).first(),
  ).toBeVisible();
  await expect(
    page.getByText(
      /constitute the complete agreement between the Parties for the Products and Services described above/,
    ).last(),
  ).toBeAttached();
});

test("keeps merged cells and repeats complex-document headers and footers", async ({
  page,
}) => {
  await page.goto("/?fixture=complex-docx");
  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 60_000,
  });

  const pages = page.locator(".docx-render-target section.docx");
  const structures = await pages.evaluateAll((sections) => ({
    colSpans: sections
      .flatMap((section) => [...section.querySelectorAll<HTMLTableCellElement>("td[colspan]")])
      .map((cell) => cell.colSpan),
    rowSpans: sections
      .flatMap((section) => [...section.querySelectorAll<HTMLTableCellElement>("td[rowspan]")])
      .map((cell) => cell.rowSpan),
    headers: sections.map(
      (section) => section.querySelector(":scope > header")?.textContent ?? "",
    ),
    footers: sections.map(
      (section) => section.querySelector(":scope > footer")?.textContent ?? "",
    ),
  }));

  expect(structures.colSpans.some((span) => span > 1)).toBe(true);
  expect(structures.rowSpans.some((span) => span > 1)).toBe(true);
  expect(
    structures.headers.filter((text) => text.includes("Master Subscription Agreement"))
      .length,
  ).toBeGreaterThan(1);
  expect(
    structures.footers.filter((text) => text.includes("Vantage Shield BV")).length,
  ).toBeGreaterThan(1);
  expect(structures.headers.length).toBeGreaterThan(1);
});

test("fits a narrow viewport without changing logical page geometry", async ({ page }) => {
  await page.setViewportSize({ width: 600, height: 900 });
  await page.goto("/?fixture=consulting-docx");
  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 30_000,
  });

  const result = await page.locator(".docx-preview-viewport").evaluate((viewport) => {
    const host = viewport.querySelector<HTMLElement>(".docx-preview");
    const root = host?.shadowRoot;
    const pageNode = root?.querySelector<HTMLElement>("section.docx");
    const sizer = root?.querySelector<HTMLElement>(".docx-scale-sizer");
    if (!pageNode || !sizer) throw new Error("DOCX layout is not ready");
    const viewportStyle = getComputedStyle(viewport);
    const viewportBounds = viewport.getBoundingClientRect();
    const sizerBounds = sizer.getBoundingClientRect();
    return {
      logicalWidth: getComputedStyle(pageNode).width,
      sizerLeft: sizerBounds.left,
      sizerRight: sizerBounds.right,
      availableLeft: viewportBounds.left + Number.parseFloat(viewportStyle.paddingLeft),
      availableRight: viewportBounds.right - Number.parseFloat(viewportStyle.paddingRight),
    };
  });

  expect(result.logicalWidth).toBe("816px");
  expect(result.sizerLeft).toBeGreaterThanOrEqual(result.availableLeft - 1);
  expect(result.sizerRight).toBeLessThanOrEqual(result.availableRight + 1);
});

test("updates review items without rereading the DOCX or replacing unaffected pages", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const original = File.prototype.arrayBuffer;
    Object.defineProperty(window, "__docxArrayBufferReads", {
      configurable: true,
      value: 0,
      writable: true,
    });
    File.prototype.arrayBuffer = function arrayBuffer() {
      const target = window as typeof window & {
        __docxArrayBufferReads: number;
      };
      target.__docxArrayBufferReads += 1;
      return original.call(this);
    };
  });
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();

  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 30_000,
  });
  const target = page.locator(".docx-render-target");
  const revision = await target.getAttribute("data-docx-live-revision");
  expect(revision).not.toBeNull();
  const firstPage = page.locator(".docx-render-target section.docx").first();
  await firstPage.evaluate((node) => {
    (
      window as typeof window & {
        __unaffectedDocxPage?: Element;
      }
    ).__unaffectedDocxPage = node;
  });
  const readsBefore = await page.evaluate(
    () =>
      (
        window as typeof window & {
          __docxArrayBufferReads: number;
        }
      ).__docxArrayBufferReads,
  );

  await page.getByTestId("item-update").click();
  await expect(target).not.toHaveAttribute(
    "data-docx-live-revision",
    revision!,
    { timeout: 30_000 },
  );
  await expect(preview).toHaveAttribute("aria-busy", "false");

  expect(
    await firstPage.evaluate(
      (node) =>
        node ===
        (
          window as typeof window & {
            __unaffectedDocxPage?: Element;
          }
        ).__unaffectedDocxPage,
    ),
  ).toBe(true);
  expect(
    await page.evaluate(
      () =>
        (
          window as typeof window & {
            __docxArrayBufferReads: number;
          }
        ).__docxArrayBufferReads,
    ),
  ).toBe(readsBefore);
});

test("retries a failed paragraph together with the next unrelated update", async ({ page }) => {
  await page.addInitScript(() => {
    const originalAppend = ShadowRoot.prototype.append;
    const testWindow = window as typeof window & {
      __failNextIncrementalAppend?: boolean;
      __incrementalAppendCount?: number;
    };
    testWindow.__failNextIncrementalAppend = false;
    testWindow.__incrementalAppendCount = 0;
    ShadowRoot.prototype.append = function append(...nodes: (Node | string)[]) {
      const incremental = nodes.some(
        (node) =>
          node instanceof HTMLElement &&
          node.classList.contains("docx-incremental-stage"),
      );
      if (incremental) {
        testWindow.__incrementalAppendCount =
          (testWindow.__incrementalAppendCount ?? 0) + 1;
        if (testWindow.__failNextIncrementalAppend) {
          testWindow.__failNextIncrementalAppend = false;
          throw new Error("Injected one-shot incremental layout failure");
        }
      }
      return originalAppend.apply(this, nodes);
    };
  });
  await page.goto("/?fixture=consulting-docx");
  await page.getByRole("button", { name: "Review" }).click();

  const preview = page.locator(".docx-preview");
  await expect(preview).toHaveAttribute("aria-busy", "false", {
    timeout: 30_000,
  });
  const target = page.locator(".docx-render-target");
  const initialRevision = await target.getAttribute("data-docx-live-revision");
  const lastGoodText = await target.textContent();

  await page.evaluate(() => {
    (
      window as typeof window & { __failNextIncrementalAppend?: boolean }
    ).__failNextIncrementalAppend = true;
  });
  await page.getByTestId("item-update").click();
  await expect(page.getByRole("alert")).toContainText(
    "Injected one-shot incremental layout failure",
  );
  await expect(target).toHaveAttribute("data-docx-live-revision", initialRevision!);
  expect(await target.textContent()).toBe(lastGoodText);

  await page.waitForTimeout(250);
  expect(
    await page.evaluate(
      () =>
        (window as typeof window & { __incrementalAppendCount?: number })
          .__incrementalAppendCount,
    ),
  ).toBe(1);

  await page.getByTestId("item-add").click();
  await expect(target).not.toHaveAttribute("data-docx-live-revision", initialRevision!, {
    timeout: 30_000,
  });
  await expect(page.getByRole("alert")).toHaveCount(0);
  await expect
    .poll(() => target.textContent())
    .toContain(
      "Company will pay all agreed professional fees after receiving a valid invoice",
    );
  await expect(
    page.locator('.docx-preview [data-docx-comment-ids*="comment-payment"]'),
  ).not.toHaveCount(0);
});
