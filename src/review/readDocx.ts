import JSZip from "jszip";

import type { DocxManifest, DocxParagraph } from "./model";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const DOCUMENT_PATH = "word/document.xml";
const MAX_DOCX_BYTES = 100 * 1024 * 1024;
const MAX_XML_BYTES = 25 * 1024 * 1024;
const MAX_FILES = 5_000;
const MAX_TEXT_NODES = 50_000;

export type ReadDocxResult = {
  manifest: DocxManifest;
  texts: Readonly<Record<string, string>>;
};

function elements(
  root: Document | Element,
  namespace: string,
  localName: string,
): Element[] {
  const list = root.getElementsByTagNameNS(namespace, localName);
  return Array.from({ length: list.length }, (_, index) => list.item(index)).filter(
    (value): value is Element => value !== null,
  );
}

function elementChildren(node: Node): Element[] {
  const result: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1) result.push(child as Element);
  }
  return result;
}

function parseXml(xml: string, label: string): Document {
  if (/<!DOCTYPE/i.test(xml)) {
    throw new Error(`${label} contains a forbidden document type declaration`);
  }
  let document: Document;
  try {
    document = new DOMParser().parseFromString(xml, "application/xml");
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error(`Could not parse ${label}: ${detail}`);
  }
  const parserError = document.getElementsByTagName("parsererror").item(0);
  if (parserError || document.documentElement?.localName === "parsererror") {
    throw new Error(`Could not parse ${label}: malformed XML`);
  }
  return document;
}

function unsupportedReason(paragraph: Element): string | undefined {
  for (
    let ancestor = paragraph.parentNode;
    ancestor?.nodeType === 1;
    ancestor = ancestor.parentNode
  ) {
    const element = ancestor as Element;
    if (element.namespaceURI === W && element.localName === "txbxContent") {
      return "Text-box paragraphs are unsupported";
    }
    if (element.namespaceURI === W && element.localName === "body") break;
  }

  for (const localName of [
    "fldChar",
    "instrText",
    "del",
    "ins",
    "moveFrom",
    "moveTo",
    "txbxContent",
  ]) {
    if (elements(paragraph, W, localName).length > 0) {
      return `Paragraph contains unsupported w:${localName} content`;
    }
  }

  for (const text of elements(paragraph, W, "t")) {
    const run = text.parentNode;
    if (
      !run ||
      run.nodeType !== 1 ||
      (run as Element).namespaceURI !== W ||
      (run as Element).localName !== "r"
    ) {
      return "Paragraph contains a text node outside a Word run";
    }
    const content = elementChildren(run).filter(
      (child) => !(child.namespaceURI === W && child.localName === "rPr"),
    );
    if (content.length !== 1 || content[0] !== text) {
      return "Paragraph contains a complex Word run";
    }
  }
  return undefined;
}

function paragraphSegments(
  paragraph: Element,
  textIds: ReadonlyMap<Element, string>,
): DocxParagraph["segments"] {
  const result: DocxParagraph["segments"] = [];
  const visit = (node: Node): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (
        element !== paragraph &&
        element.namespaceURI === W &&
        element.localName === "p"
      ) {
        continue;
      }
      if (element.namespaceURI === W && element.localName === "t") {
        const textId = textIds.get(element);
        if (textId) result.push({ textId });
      } else if (element.namespaceURI === W && element.localName === "tab") {
        result.push({ literal: "\t" });
      } else if (
        element.namespaceURI === W &&
        (element.localName === "br" || element.localName === "cr")
      ) {
        result.push({ literal: "\n" });
      } else {
        visit(element);
      }
    }
  };
  visit(paragraph);
  return result;
}

function nativeParagraphId(paragraph: Element): string | undefined {
  return (
    paragraph.getAttributeNS(W14, "paraId") ||
    paragraph.getAttribute("w14:paraId") ||
    undefined
  );
}

export async function readDocx(bytes: Uint8Array): Promise<ReadDocxResult> {
  if (bytes.byteLength === 0) throw new Error("DOCX input is empty");
  if (bytes.byteLength > MAX_DOCX_BYTES) {
    throw new Error(`DOCX input exceeds ${MAX_DOCX_BYTES} bytes`);
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error(`Could not open DOCX package: ${detail}`);
  }
  if (Object.keys(zip.files).length > MAX_FILES) {
    throw new Error(`DOCX package contains more than ${MAX_FILES} files`);
  }
  const documentFile = zip.file(DOCUMENT_PATH);
  if (!documentFile) throw new Error(`DOCX package is missing ${DOCUMENT_PATH}`);
  const xml = await documentFile.async("string");
  if (xml.length > MAX_XML_BYTES) {
    throw new Error(`${DOCUMENT_PATH} exceeds ${MAX_XML_BYTES} bytes`);
  }

  const document = parseXml(xml, DOCUMENT_PATH);
  const body = elements(document, W, "body")[0];
  if (!body) throw new Error(`${DOCUMENT_PATH} has no w:body`);
  const textElements = elements(document, W, "t");
  if (textElements.length > MAX_TEXT_NODES) {
    throw new Error(`${DOCUMENT_PATH} contains more than ${MAX_TEXT_NODES} text nodes`);
  }

  const textIds = new Map<Element, string>();
  const texts: Record<string, string> = {};
  const manifestTexts = textElements.map((element, index) => {
    const id = `t${index}`;
    textIds.set(element, id);
    texts[id] = element.textContent ?? "";
    return { id, index };
  });
  const paragraphs = elements(body, W, "p").map((element, index) => {
    const reason = unsupportedReason(element);
    const nativeId = nativeParagraphId(element);
    return {
      id: `p${index}`,
      segments: paragraphSegments(element, textIds),
      ...(nativeId ? { nativeId } : {}),
      ...(reason ? { editable: false as const, reason } : {}),
    };
  });

  return {
    manifest: { schemaVersion: 1, paragraphs, texts: manifestTexts },
    texts,
  };
}
