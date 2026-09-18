import { cpSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const exampleRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const repositoryRoot = resolve(exampleRoot, "..");
const publicRoot = resolve(exampleRoot, "public");
const pdfRoot = resolve(publicRoot, "pdfjs");

mkdirSync(publicRoot, { recursive: true });
mkdirSync(pdfRoot, { recursive: true });
cpSync(resolve(repositoryRoot, "tests/preview/fixtures"), resolve(publicRoot, "fixtures"), {
  recursive: true,
  force: true,
});
cpSync(
  resolve(repositoryRoot, "node_modules/pdfjs-dist/build/pdf.worker.min.mjs"),
  resolve(pdfRoot, "pdf.worker.min.mjs"),
  { force: true },
);
for (const directory of ["cmaps", "standard_fonts", "wasm"]) {
  cpSync(
    resolve(repositoryRoot, `node_modules/pdfjs-dist/${directory}`),
    resolve(pdfRoot, directory),
    { recursive: true, force: true },
  );
}
