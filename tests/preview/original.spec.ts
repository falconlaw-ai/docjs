import { expect, test } from "@playwright/test";

test("renders DOCX, PDF, and Markdown originals with their native adapters", async ({
  page,
}) => {
  await page.goto("/?fixture=consulting-docx");
  const docx = page.locator(".docx-preview");
  await expect(docx).toHaveAttribute("aria-busy", "false", { timeout: 30_000 });
  await expect(
    page.locator(".docx-render-target section.docx").first(),
  ).toBeVisible();

  await page.getByLabel("Example document").selectOption("consulting-pdf");
  const pdf = page.locator(".pdf-preview");
  await expect(pdf).toHaveAttribute("aria-busy", "false", { timeout: 30_000 });
  await expect(pdf.locator(".pdf-preview__page")).toHaveCount(5);
  await expect(pdf.locator("canvas").first()).toBeVisible();

  await page.getByLabel("Example document").selectOption("consulting-markdown");
  const markdown = page.locator(".markdown-preview");
  await expect(markdown).toHaveAttribute("aria-busy", "false", {
    timeout: 30_000,
  });
  await expect(
    markdown.getByRole("heading", { name: "CONSULTING AGREEMENT" }),
  ).toBeVisible();
});

test("fits original pages without changing their logical geometry", async ({
  page,
}) => {
  await page.goto("/?fixture=consulting-docx");
  const docx = page.locator(".docx-preview");
  await expect(docx).toHaveAttribute("aria-busy", "false", { timeout: 30_000 });
  const pageNode = page.locator(".docx-render-target section.docx").first();
  const logicalSize = await pageNode.evaluate((element) => {
    const style = getComputedStyle(element);
    return { height: style.height, width: style.width };
  });

  await page.getByLabel("Zoom percentage").fill("75");
  await page.getByLabel("Zoom percentage").press("Enter");
  await expect
    .poll(() =>
      page.locator(".docx-scale-layer").evaluate((element) => element.style.transform),
    )
    .toBe("scale(0.75)");
  await expect
    .poll(() =>
      pageNode.evaluate((element) => {
        const style = getComputedStyle(element);
        return { height: style.height, width: style.width };
      }),
    )
    .toEqual(logicalSize);

  await page.getByRole("button", { name: "Fit width" }).click();
  await expect(page.getByRole("button", { name: "Fit width" })).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  await expect
    .poll(() =>
      pageNode.evaluate((element) => {
        const style = getComputedStyle(element);
        return { height: style.height, width: style.width };
      }),
    )
    .toEqual(logicalSize);
});

test("does not fetch an external image supplied by a Markdown document", async ({
  page,
}) => {
  const externalRequests: string[] = [];
  page.on("request", (request) => {
    if (request.url().startsWith("https://example.invalid/")) {
      externalRequests.push(request.url());
    }
  });

  await page.goto("/?fixture=consulting-markdown");
  await page.getByLabel("Open a document").setInputFiles({
    name: "external-image.md",
    mimeType: "text/markdown",
    buffer: Buffer.from(
      "# External image\n\n![Tracking pixel](https://example.invalid/tracker.png)",
    ),
  });

  await expect(page.getByRole("heading", { name: "External image" })).toBeVisible();
  await expect(page.getByRole("note")).toHaveText("[Image: Tracking pixel]");
  await page.waitForTimeout(100);
  expect(externalRequests).toEqual([]);
});
