import { readFile } from "node:fs/promises";
import typescript from "@rollup/plugin-typescript";
import terser from "@rollup/plugin-terser";

const banner = `/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */`;

const output = { banner, sourcemap: true };
const umdOutput = {
  ...output,
  name: "docx",
  file: "dist/docx-preview.js",
  format: "umd",
  globals: { jszip: "JSZip" },
};

function reactCss() {
  const styles = new Map();
  return {
    name: "react-css",
    async load(id) {
      if (!id.endsWith(".css")) return null;
      const source = await readFile(id, "utf8");
      styles.set(id, source);
      return `export default ${JSON.stringify(source)};`;
    },
    generateBundle() {
      const stylesheet = [...styles]
        .filter(([id]) => !id.endsWith("/docx-preview.css"))
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([, source]) => source)
        .join("\n");
      this.emitFile({ type: "asset", fileName: "react/styles.css", source: stylesheet });
    },
  };
}

const external = (id) =>
  !id.startsWith(".") && !id.startsWith("/") && !id.startsWith("\0");

export default (args) => {
  const production = args.environment === "BUILD:production";
  const engineOutputs = production
    ? [
        umdOutput,
        { ...umdOutput, file: "dist/docx-preview.min.js", plugins: [terser()] },
        { ...output, file: "dist/docx-preview.mjs", format: "es" },
        {
          ...output,
          file: "dist/docx-preview.min.mjs",
          format: "es",
          plugins: [terser()],
        },
      ]
    : [umdOutput];

  return [
    {
      input: "src/docx-preview.ts",
      output: engineOutputs,
      external,
      plugins: [typescript({ tsconfig: "./tsconfig.json" })],
    },
    {
      input: "src/react/index.ts",
      output: {
        ...output,
        dir: "dist",
        format: "es",
        entryFileNames: "react/index.mjs",
        chunkFileNames: "react/chunks/[name]-[hash].mjs",
      },
      external,
      plugins: [reactCss(), typescript({ tsconfig: "./tsconfig.react.json" })],
    },
  ];
};
