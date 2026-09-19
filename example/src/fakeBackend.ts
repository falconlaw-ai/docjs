import type { DocumentFormat, PreviewDocument } from "docx-preview/react";

export type FixtureId =
  | "consulting-docx"
  | "consulting-pdf"
  | "consulting-markdown"
  | "complex-docx"
  | "complex-pdf"
  | "complex-markdown"
  | "preparation-error";

type Fixture = {
  id: FixtureId;
  label: string;
  original: string;
  format: DocumentFormat;
  prepared?: string;
  failFirstPreparation?: boolean;
};

export const fixtures: readonly Fixture[] = [
  {
    id: "consulting-docx",
    label: "Consulting agreement · DOCX",
    original: "/fixtures/consulting_agreement.docx",
    format: "docx",
  },
  {
    id: "consulting-pdf",
    label: "Consulting agreement · PDF",
    original: "/fixtures/consulting_agreement.pdf",
    prepared: "/fixtures/consulting_agreement.docx",
    format: "pdf",
  },
  {
    id: "consulting-markdown",
    label: "Consulting agreement · Markdown",
    original: "/fixtures/consulting_agreement.md",
    prepared: "/fixtures/consulting_agreement.docx",
    format: "md",
  },
  {
    id: "complex-docx",
    label: "Complex MSA · DOCX",
    original: "/fixtures/msa_vantage_shield_complex.docx",
    format: "docx",
  },
  {
    id: "complex-pdf",
    label: "Complex MSA · PDF",
    original: "/fixtures/msa_vantage_shield_complex.pdf",
    prepared: "/fixtures/msa_vantage_shield_complex.docx",
    format: "pdf",
  },
  {
    id: "complex-markdown",
    label: "Complex MSA · Markdown",
    original: "/fixtures/msa_vantage_shield_complex.md",
    prepared: "/fixtures/msa_vantage_shield_complex.docx",
    format: "md",
  },
  {
    id: "preparation-error",
    label: "Preparation failure then retry",
    original: "/fixtures/consulting_agreement.md",
    prepared: "/fixtures/consulting_agreement.docx",
    format: "md",
    failFirstPreparation: true,
  },
];

const mimeTypes: Record<DocumentFormat, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  md: "text/markdown",
};

const preparationAttempts = new Map<string, number>();

function fileName(url: string) {
  return url.slice(url.lastIndexOf("/") + 1);
}

async function fetchFile(url: string, format: DocumentFormat, signal?: AbortSignal) {
  const response = await fetch(url, { signal });
  if (!response.ok) throw new Error(`Fixture request failed with ${response.status}.`);
  return new File([await response.blob()], fileName(url), { type: mimeTypes[format] });
}

export async function loadFixture(id: FixtureId, signal?: AbortSignal): Promise<PreviewDocument> {
  const fixture = fixtures.find((candidate) => candidate.id === id);
  if (!fixture) throw new Error(`Unknown fixture: ${id}`);
  return {
    id: fixture.id,
    revision: "fixture-v1",
    format: fixture.format,
    file: await fetchFile(fixture.original, fixture.format, signal),
  };
}

function wait(delay: number, signal: AbortSignal) {
  return new Promise<void>((resolve, reject) => {
    const timeout = window.setTimeout(resolve, delay);
    signal.addEventListener(
      "abort",
      () => {
        window.clearTimeout(timeout);
        reject(new DOMException("Preparation was canceled.", "AbortError"));
      },
      { once: true },
    );
  });
}

export async function prepareFixture(
  original: PreviewDocument,
  options: { delay: number; signal: AbortSignal },
) {
  const fixture = fixtures.find((candidate) => candidate.id === original.id);
  if (!fixture?.prepared) throw new Error("This original has no deterministic preparation mapping.");
  await wait(options.delay, options.signal);
  const attemptKey = `${original.id}:${original.revision}`;
  const attempt = (preparationAttempts.get(attemptKey) ?? 0) + 1;
  preparationAttempts.set(attemptKey, attempt);
  if (fixture.failFirstPreparation && attempt === 1) {
    throw new Error("The example backend rejected this preparation. Retry to continue.");
  }
  return fetchFile(fixture.prepared, "docx", options.signal);
}
