const assert = require("node:assert/strict");
const { readFile } = require("node:fs/promises");
const { pathToFileURL } = require("node:url");
const test = require("node:test");

test("keeps the CommonJS engine entry and ESM React entry loadable", async () => {
  const engine = require("../../dist/docx-preview.js");
  assert.equal(typeof engine.renderAsync, "function");

  const react = await import(pathToFileURL(`${process.cwd()}/dist/react/index.mjs`));
  assert.equal(typeof react.DocumentPreview, "function");
});

test("exports declarations and the packaged stylesheet", async () => {
  const files = await Promise.all([
    readFile("dist/docx-preview.d.ts", "utf8"),
    readFile("dist/react/index.d.ts", "utf8"),
    readFile("dist/react/styles.css", "utf8"),
  ]);
  assert.ok(files.every((source) => source.length > 0));
});
