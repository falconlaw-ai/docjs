import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const IGNORED_CACHED_PAGE_BREAK = "ignoredLastRenderedPageBreak";

type ModelNode = {
  type: string;
  text?: string;
  break?: string;
  src?: string;
  className?: string;
  cssStyle?: Record<string, string>;
  children?: ModelNode[];
  sectionProps?: SectionProps;
};

type HeaderFooterRef = { id: string; type: string };
type SectionProps = {
  type?: string;
  titlePage?: boolean;
  headerRefs?: HeaderFooterRef[];
  footerRefs?: HeaderFooterRef[];
};

type ModelDocument = {
  documentPart: {
    body: ModelNode & { props?: SectionProps };
  };
};

type Position = {
  relativeFrom: string;
  align?: string;
  offset?: number;
};

type AnchorGeometry = {
  className: string;
  behindText: boolean;
  width: number;
  height: number;
  horizontal: Position;
  vertical: Position;
  wrap: "none" | "square" | "tight" | "topAndBottom" | "through";
  distances: { top: number; right: number; bottom: number; left: number };
};

export type DocxGeometryPlan = {
  anchors: AnchorGeometry[];
  paragraphBorders?: Array<{
    className: string;
    spacing: Partial<Record<"left" | "right" | "top" | "bottom", number>>;
  }>;
};

type XmlDrawing = Omit<AnchorGeometry, "className"> & {
  relationshipId?: string;
  anchor: boolean;
};

const emuToPoints = (value: string | null): number =>
  Number.parseInt(value ?? "0", 10) / 12_700;

function child(element: Element, namespace: string, localName: string) {
  return [...element.children].find(
    (candidate) =>
      candidate.namespaceURI === namespace && candidate.localName === localName,
  );
}

function position(element: Element | undefined): Position {
  if (!element) return { relativeFrom: "page" };
  const align = child(element, WP, "align")?.textContent ?? undefined;
  const offsetText = child(element, WP, "posOffset")?.textContent;
  return {
    relativeFrom: element.getAttribute("relativeFrom") ?? "page",
    align,
    offset: offsetText == null ? undefined : emuToPoints(offsetText),
  };
}

function parseDrawing(element: Element): XmlDrawing | undefined {
  const wrapper = [...element.children].find(
    (candidate) =>
      candidate.namespaceURI === WP &&
      (candidate.localName === "anchor" || candidate.localName === "inline"),
  );
  if (!wrapper) return undefined;
  const extent = child(wrapper, WP, "extent");
  const blip = wrapper.getElementsByTagNameNS(A, "blip")[0];
  const wrap = [...wrapper.children].find(
    (candidate) =>
      candidate.namespaceURI === WP && candidate.localName.startsWith("wrap"),
  );
  const wrapName = wrap?.localName;
  const wrapType =
    wrapName === "wrapSquare"
      ? "square"
      : wrapName === "wrapTight"
        ? "tight"
        : wrapName === "wrapTopAndBottom"
          ? "topAndBottom"
          : wrapName === "wrapThrough"
            ? "through"
            : "none";
  return {
    relationshipId: blip?.getAttributeNS(R, "embed") ?? undefined,
    anchor: wrapper.localName === "anchor",
    behindText: ["1", "true"].includes(
      wrapper.getAttribute("behindDoc")?.toLowerCase() ?? "",
    ),
    width: emuToPoints(extent?.getAttribute("cx") ?? null),
    height: emuToPoints(extent?.getAttribute("cy") ?? null),
    horizontal: position(child(wrapper, WP, "positionH")),
    vertical: position(child(wrapper, WP, "positionV")),
    wrap: wrapType,
    distances: {
      top: emuToPoints(wrapper.getAttribute("distT")),
      right: emuToPoints(wrapper.getAttribute("distR")),
      bottom: emuToPoints(wrapper.getAttribute("distB")),
      left: emuToPoints(wrapper.getAttribute("distL")),
    },
  };
}

function visit(node: ModelNode, visitor: (node: ModelNode) => void): void {
  visitor(node);
  node.children?.forEach((item) => visit(item, visitor));
}

function containsPageBreak(node: ModelNode): boolean {
  let found = false;
  visit(node, (item) => {
    if (item.type === "break" && item.break === "page") found = true;
  });
  return found;
}

function neutralizeCachedPageBreaks(node: ModelNode): void {
  visit(node, (item) => {
    if (item.type === "break" && item.break === "lastRenderedPageBreak") {
      // Keep the marker node and its run in place for stable live identities,
      // but don't let a cached Word layout boundary split browser flow.
      item.break = IGNORED_CACHED_PAGE_BREAK;
    }
  });
}

const EMPTY_PARAGRAPH_NODE_TYPES = new Set([
  "paragraph",
  "run",
  "text",
  "deletedText",
  "hyperlink",
  "smartTag",
  "inserted",
  "deleted",
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
]);

function isVisiblyEmptyParagraph(paragraph: ModelNode): boolean {
  let empty = true;
  visit(paragraph, (node) => {
    const ignoredCachedBreak =
      node.type === "break" && node.break === IGNORED_CACHED_PAGE_BREAK;
    if (
      (node.text?.length ?? 0) > 0 ||
      (!EMPTY_PARAGRAPH_NODE_TYPES.has(node.type) && !ignoredCachedBreak)
    ) {
      empty = false;
    }
  });
  return empty;
}

function collapseContinuousSectionCarrier(
  paragraph: ModelNode,
  nextSection: SectionProps | undefined,
): void {
  if (
    paragraph.type !== "paragraph" ||
    !paragraph.sectionProps ||
    !["continuous", "nextColumn"].includes(nextSection?.type ?? "") ||
    !isVisiblyEmptyParagraph(paragraph)
  ) {
    return;
  }
  // The paragraph carries the preceding section's metadata, but Word does not
  // give an empty same-page section boundary its inherited line box or spacing.
  paragraph.className =
    `${paragraph.className ?? ""} docx-section-carrier`.trim();
  paragraph.cssStyle ??= {};
  paragraph.cssStyle.height = "0";
  paragraph.cssStyle["line-height"] = "0";
  paragraph.cssStyle["min-height"] = "0";
  paragraph.cssStyle["margin-top"] = "0";
  paragraph.cssStyle["margin-bottom"] = "0";
}

function addSectionPageBreak(
  paragraph: ModelNode,
  props: SectionProps | undefined,
): boolean {
  if (!props || props.type === "continuous" || props.type === "nextColumn")
    return false;
  if (!containsPageBreak(paragraph)) {
    paragraph.children ??= [];
    paragraph.children.push({
      type: "run",
      children: [{ type: "break", break: "page" }],
    });
  }
  return true;
}

function fillLinkedRefs(
  own: HeaderFooterRef[] | undefined,
  inherited: Map<string, HeaderFooterRef>,
): HeaderFooterRef[] | undefined {
  const refs = new Map(inherited);
  own?.forEach((ref) => refs.set(ref.type, ref));
  return refs.size > 0 ? [...refs.values()] : undefined;
}

function repairSections(body: ModelNode & { props?: SectionProps }): void {
  neutralizeCachedPageBreaks(body);
  const sections = (body.children ?? []).flatMap((node) =>
    node.sectionProps ? [node.sectionProps] : [],
  );
  if (body.props) sections.push(body.props);
  let sectionIndex = 0;
  for (const node of body.children ?? []) {
    // sectPr closes a section, but its type describes how that section STARTS
    // relative to its predecessor. The boundary here belongs to the next one.
    const nextSection = node.sectionProps
      ? sections[++sectionIndex]
      : undefined;
    collapseContinuousSectionCarrier(node, nextSection);
    if (node.sectionProps) addSectionPageBreak(node, nextSection);
  }

  let headers = new Map<string, HeaderFooterRef>();
  let footers = new Map<string, HeaderFooterRef>();
  sections.forEach((section, index) => {
    section.headerRefs = fillLinkedRefs(section.headerRefs, headers);
    section.footerRefs = fillLinkedRefs(section.footerRefs, footers);
    if (
      section.titlePage &&
      !section.headerRefs?.some((reference) => reference.type === "first")
    ) {
      section.headerRefs ??= [];
      section.headerRefs.push({
        id: `falcon-missing-first-header-${index}`,
        type: "first",
      });
    }
    if (
      section.titlePage &&
      !section.footerRefs?.some((reference) => reference.type === "first")
    ) {
      section.footerRefs ??= [];
      section.footerRefs.push({
        id: `falcon-missing-first-footer-${index}`,
        type: "first",
      });
    }
    headers = new Map(section.headerRefs?.map((ref) => [ref.type, ref]));
    footers = new Map(section.footerRefs?.map((ref) => [ref.type, ref]));
  });
}

function annotateAnchors(
  body: ModelNode,
  xmlDrawings: XmlDrawing[],
): AnchorGeometry[] {
  const modelDrawings: ModelNode[] = [];
  visit(body, (node) => {
    if (node.type === "drawing") modelDrawings.push(node);
  });
  const anchors: AnchorGeometry[] = [];
  const length = Math.min(modelDrawings.length, xmlDrawings.length);
  for (let index = 0; index < length; index++) {
    const geometry = xmlDrawings[index];
    if (!geometry.anchor) continue;
    const model = modelDrawings[index];
    const modelRelationship = model.children?.find(
      (item) => item.type === "image",
    )?.src;
    if (
      geometry.relationshipId &&
      modelRelationship &&
      geometry.relationshipId !== modelRelationship
    )
      continue;
    const className = `docx-geometry-anchor-${anchors.length}`;
    model.className = `${model.className ?? ""} ${className}`.trim();
    anchors.push({ ...geometry, className });
  }
  return anchors;
}

function annotateParagraphBorders(xml: Document, body: ModelNode) {
  const paragraphs = new Map<string, ModelNode[]>();
  const textOf = (node: ModelNode): string =>
    node.text ?? (node.children ?? []).map(textOf).join("");
  visit(body, (node) => {
    if (node.type !== "paragraph") return;
    const text = textOf(node);
    paragraphs.set(text, [...(paragraphs.get(text) ?? []), node]);
  });
  const sourceParagraphs = new Map<string, Element[]>();
  for (const paragraph of xml.getElementsByTagNameNS(W, "p")) {
    const text = [...paragraph.getElementsByTagNameNS(W, "*")]
      .filter((node) => node.localName === "t" || node.localName === "delText")
      .map((node) => node.textContent ?? "")
      .join("");
    sourceParagraphs.set(text, [...(sourceParagraphs.get(text) ?? []), paragraph]);
  }
  const result: NonNullable<DocxGeometryPlan["paragraphBorders"]> = [];
  for (const [text, sources] of sourceParagraphs) {
    const models = paragraphs.get(text);
    // Unsupported source content can be absent from the parsed model. Don't
    // assign formatting to a different occurrence when the mapping is ambiguous.
    if (!models || models.length !== sources.length) continue;
    sources.forEach((source, index) => {
      const properties = child(source, W, "pPr");
      const borders = properties && child(properties, W, "pBdr");
      if (!borders) return;
      const spacing: (typeof result)[number]["spacing"] = {};
      for (const side of ["left", "right", "top", "bottom"] as const) {
        const border = child(borders, W, side);
        const points = Number(border?.getAttributeNS(W, "space"));
        if (Number.isFinite(points) && points > 0) spacing[side] = points;
      }
      if (!Object.keys(spacing).length) return;
      const className = `docx-paragraph-border-${result.length}`;
      const model = models[index];
      model.className = `${model.className ?? ""} ${className}`.trim();
      result.push({ className, spacing });
    });
  }
  return result;
}

/** Repair geometry metadata in docx-preview's disposable parsed model. */
export async function prepareDocxGeometry(
  source: ArrayBuffer,
  document: ModelDocument,
): Promise<DocxGeometryPlan> {
  const zip = await JSZip.loadAsync(source);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  repairSections(document.documentPart.body);
  if (!documentXml) return { anchors: [] };
  const xml = new DOMParser().parseFromString(documentXml, "application/xml");
  if (xml.querySelector("parsererror")) return { anchors: [] };
  const xmlDrawings = [...xml.getElementsByTagNameNS(W, "drawing")]
    .map(parseDrawing)
    .filter((item): item is XmlDrawing => item != null);
  return {
    anchors: annotateAnchors(document.documentPart.body, xmlDrawings),
    paragraphBorders: annotateParagraphBorders(xml, document.documentPart.body),
  };
}

function alignedPosition(position: Position, size: number): string | undefined {
  if (position.align === "center") return `calc(50% - ${size / 2}pt)`;
  if (position.align === "right" || position.align === "bottom")
    return `calc(100% - ${size}pt)`;
  if (position.align === "left" || position.align === "top") return "0";
  return position.offset == null ? undefined : `${position.offset}pt`;
}

/** Restore anchor semantics that docx-preview 0.4.0 drops while parsing. */
export function applyDocxGeometry(
  container: HTMLElement,
  plan: DocxGeometryPlan,
): void {
  for (const border of plan.paragraphBorders ?? []) {
    for (const paragraph of container.querySelectorAll<HTMLElement>(
      `p.${border.className}`,
    )) {
      const computed = getComputedStyle(paragraph);
      for (const [side, points] of Object.entries(border.spacing)) {
        const width = computed.getPropertyValue(`border-${side}-width`);
        if (!Number.parseFloat(width)) continue;
        paragraph.style.setProperty(`padding-${side}`, `${points}pt`);
        if (side === "left" || side === "right") {
          // Word places the border outside the paragraph's text indent. CSS
          // padding alone would shift the text and narrow its wrapping width.
          const margin = computed.getPropertyValue(`margin-${side}`);
          paragraph.style.setProperty(
            `margin-${side}`,
            `calc(${margin} - ${points}pt - ${width})`,
          );
        }
      }
    }
  }
  for (const layer of container.querySelectorAll<HTMLElement>(
    "section.docx > article, section.docx > header, section.docx > footer",
  )) {
    layer.style.position = "relative";
    layer.style.zIndex = "1";
  }
  // A non-floating Word table starts below a preceding floating drawing. If it
  // participates in CSS float wrapping, auto layout permanently narrows its
  // stored grid before the paginator can freeze the measured columns.
  for (const table of container.querySelectorAll<HTMLElement>(
    "section.docx > article > table",
  )) {
    if (!table.style.cssFloat || table.style.cssFloat === "none")
      table.style.clear = "both";
  }
  for (const anchor of plan.anchors) {
    const wrapper = container.querySelector<HTMLElement>(`.${anchor.className}`);
    if (!wrapper) continue;
    if (anchor.behindText) {
      const page = wrapper.closest<HTMLElement>("section.docx");
      if (!page) continue;
      wrapper.dataset.docxBehindText = "true";
      wrapper.dataset.docxPageDecoration = "true";
      wrapper.style.position = "absolute";
      wrapper.style.display = "block";
      wrapper.style.width = `${anchor.width}pt`;
      wrapper.style.height = `${anchor.height}pt`;
      wrapper.style.left = alignedPosition(anchor.horizontal, anchor.width) ?? "0";
      wrapper.style.top = alignedPosition(anchor.vertical, anchor.height) ?? "0";
      wrapper.style.zIndex = "0";
      wrapper.style.pointerEvents = "none";
      page.prepend(wrapper);
      continue;
    }
    if (
      anchor.wrap === "square" ||
      anchor.wrap === "tight" ||
      anchor.wrap === "through"
    ) {
      wrapper.dataset.docxFloatingAnchor = "true";
      wrapper.style.marginTop = `${anchor.distances.top}pt`;
      wrapper.style.marginRight = `${anchor.distances.right}pt`;
      wrapper.style.marginBottom = `${anchor.distances.bottom}pt`;
      wrapper.style.marginLeft = `${anchor.distances.left}pt`;
    }
  }
}
