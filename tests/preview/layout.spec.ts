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
