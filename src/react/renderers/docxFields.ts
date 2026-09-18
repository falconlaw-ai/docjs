import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
type FieldKind = "page" | "pages";
type FieldTokens = Record<FieldKind, string>;
type ModelNode = {
  type: string;
  text?: string;
  className?: string;
  children?: ModelNode[];
};

function kind(instruction: string): FieldKind | undefined {
  const name = instruction.trim().split(/\s+/)[0]?.toUpperCase();
  return name === "PAGE" ? "page" : name === "NUMPAGES" ? "pages" : undefined;
}

function placeholder(
  owner: XMLDocument,
  token: string,
  styledRun: Element | undefined,
): Element {
  const run = owner.createElementNS(W, "w:r");
  const properties =
    styledRun &&
    [...styledRun.children].find(
      (child) => child.namespaceURI === W && child.localName === "rPr",
    );
  if (properties) run.append(properties.cloneNode(true));
  const text = owner.createElementNS(W, "w:t");
  text.textContent = token;
  run.append(text);
  return run;
}

function outsideFieldRun(
  run: Element,
  boundary: Element,
  before: boolean,
): Element | undefined {
  const children = [...run.children];
  const edge = children.indexOf(boundary);
  const content = (
    before ? children.slice(0, edge) : children.slice(edge + 1)
  ).filter((child) => child.localName !== "rPr");
  if (!content.length) return;
  const copy = run.cloneNode(false) as Element;
  const properties = children.find((child) => child.localName === "rPr");
  if (properties) copy.append(properties.cloneNode(true));
  content.forEach((child) => copy.append(child.cloneNode(true)));
  return copy;
}

/** Normalize supported fields in a disposable preview package, never the source/export. */
export async function preparePageFields(source: ArrayBuffer) {
  const tokens: FieldTokens = {
    page: `FALCON-PREVIEW-PAGE-${crypto.randomUUID()}`,
    pages: `FALCON-PREVIEW-PAGES-${crypto.randomUUID()}`,
  };
  const zip = await JSZip.loadAsync(source);
  let changed = false;
  for (const file of Object.values(zip.files)) {
    if (!/^word\/(document|header[^/]*|footer[^/]*)\.xml$/.test(file.name))
      continue;
    const xml = new DOMParser().parseFromString(
      await file.async("string"),
      "application/xml",
    );
    if (xml.querySelector("parsererror")) continue;
    let updated = false;
    for (const field of [...xml.getElementsByTagNameNS(W, "fldSimple")]) {
      const fieldKind = kind(field.getAttributeNS(W, "instr") ?? "");
      if (!fieldKind) continue;
      const run = field.getElementsByTagNameNS(W, "r")[0];
      field.replaceWith(placeholder(xml, tokens[fieldKind], run));
      updated = true;
    }
    // Complex fields span sibling runs. Preserve unsupported/nested field structures.
    for (const paragraph of [...xml.getElementsByTagNameNS(W, "p")]) {
      const stack: {
        start: Element;
        startControl: Element;
        instruction: string;
        nested: boolean;
        result?: Element;
      }[] = [];
      for (const run of [...paragraph.children]) {
        if (run.namespaceURI !== W || run.localName !== "r") continue;
        for (const child of [...run.children]) {
          if (child.namespaceURI !== W) continue;
          const current = stack.at(-1);
          if (child.localName === "instrText" && current)
            current.instruction += child.textContent ?? "";
          if (child.localName === "t" && current && !current.result)
            current.result = run;
          if (child.localName !== "fldChar") continue;
          const phase = child.getAttributeNS(W, "fldCharType");
          if (phase === "begin") {
            if (current) current.nested = true;
            stack.push({
              start: run,
              startControl: child,
              instruction: "",
              nested: false,
            });
          } else if (phase === "end" && current) {
            stack.pop();
            const fieldKind = kind(current.instruction);
            if (!fieldKind || current.nested || stack.length) continue;
            const replacement = placeholder(
              xml,
              tokens[fieldKind],
              current.result ?? current.start,
            );
            const prefix = outsideFieldRun(
              current.start,
              current.startControl,
              true,
            );
            const suffix = outsideFieldRun(run, child, false);
            if (prefix) paragraph.insertBefore(prefix, current.start);
            paragraph.insertBefore(replacement, current.start);
            if (suffix) run.after(suffix);
            let node: Element | null = current.start;
            while (node) {
              const next: Element | null = node.nextElementSibling;
              node.remove();
              if (node === run) break;
              node = next;
            }
            updated = true;
          }
        }
      }
    }
    if (updated) {
      zip.file(file.name, new XMLSerializer().serializeToString(xml));
      changed = true;
    }
  }
  return {
    buffer: changed ? await zip.generateAsync({ type: "arraybuffer" }) : source,
    tokens,
  };
}

export function markPageFields(
  document: { parts: { body?: ModelNode; rootElement?: ModelNode }[] },
  tokens: FieldTokens,
): void {
  function visit(node: ModelNode) {
    for (const child of node.children ?? []) {
      const fieldKind =
        child.type === "text"
          ? (Object.keys(tokens) as FieldKind[]).find(
              (key) => tokens[key] === child.text,
            )
          : undefined;
      if (fieldKind) {
        node.className =
          `${node.className ?? ""} docx-field-${fieldKind}`.trim();
        child.text = "1";
      } else visit(child);
    }
  }
  for (const part of document.parts) {
    const root = part.body ?? part.rootElement;
    if (root) visit(root);
  }
}

/** Use the actual preview pages; cached Word page numbers can be stale after edits. */
export function updatePageFields(container: HTMLElement): void {
  const pages = [...container.querySelectorAll(".docx-wrapper > section.docx")];
  pages.forEach((page, index) => {
    page.querySelectorAll(".docx-field-page").forEach((field) => {
      field.textContent = String(index + 1);
    });
    page.querySelectorAll(".docx-field-pages").forEach((field) => {
      field.textContent = String(pages.length);
    });
  });
}
