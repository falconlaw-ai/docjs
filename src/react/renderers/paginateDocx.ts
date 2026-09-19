/** Browser-measured pagination of docx-preview's DOM, without changing the DOCX. */
type ParagraphHints = {
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
};
type LayoutNode = ParagraphHints & {
  type: string;
  styleName?: string;
  className?: string;
  cssStyle?: Record<string, string>;
  children?: LayoutNode[];
  isHeader?: boolean | null;
};
type LayoutStyle = {
  id: string;
  basedOn?: string;
  target?: string;
  isDefault?: boolean;
  paragraphProps?: ParagraphHints;
};
type LayoutDocument = {
  documentPart: { body: LayoutNode };
  stylesPart?: { styles: LayoutStyle[] };
};

// docx-preview 0.4 parses these properties, but doesn't emit them as layout CSS.
export function applyPaginationHints(document: LayoutDocument): void {
  const styles = new Map(document.stylesPart?.styles.map((s) => [s.id, s]));
  const defaultStyle = document.stylesPart?.styles.find(
    (s) => s.target === "p" && s.isDefault,
  );
  function inherited(
    id: string | undefined,
    seen = new Set<string>(),
  ): ParagraphHints {
    if (!id || seen.has(id)) return {};
    seen.add(id);
    const style = styles.get(id);
    return { ...inherited(style?.basedOn, seen), ...style?.paragraphProps };
  }
  function visit(node: LayoutNode) {
    if (node.type === "paragraph") {
      const hints = {
        ...inherited(defaultStyle?.id),
        ...inherited(node.styleName),
      };
      node.cssStyle ??= {};
      if (node.pageBreakBefore ?? hints.pageBreakBefore)
        node.cssStyle["break-before"] = "page";
      if (node.keepNext ?? hints.keepNext)
        node.cssStyle["break-after"] = "avoid";
      if (node.keepLines ?? hints.keepLines)
        node.cssStyle["break-inside"] = "avoid";
    }
    // The parser returns null for <w:tblHeader/>, an enabled OOXML boolean.
    if (node.isHeader === true || node.isHeader === null) {
      // Its DOM builder uses Object.assign(style), which drops CSS variables.
      node.className = `${node.className ?? ""} docx-repeat-header`.trim();
    }
    node.children?.forEach(visit);
  }
  visit(document.documentPart.body);
}

const px = (value: string) => Number.parseFloat(value) || 0;
function bottom(element: Element): number {
  let result = element.getBoundingClientRect().bottom;
  // Anchored drawings may paint outside a zero-height paragraph wrapper.
  for (const drawing of element.querySelectorAll("img, svg")) {
    if (
      !drawing.closest(".docx-comment-popover") &&
      !drawing.closest("[data-docx-behind-text]")
    ) {
      result = Math.max(result, drawing.getBoundingClientRect().bottom);
    }
  }
  return result;
}
const fits = (element: Element, limit: number) =>
  bottom(element) <= limit + 0.5;

function fitOversized(block: HTMLElement, limit: number): void {
  const top = block.getBoundingClientRect().top;
  const available = limit - top;
  if (available <= 0) {
    throw new Error(
      "The document's header and margins leave no space for page content.",
    );
  }
  const extent = bottom(block) - top;
  // Reserve painted height so a floating image cannot overlap the next block.
  if (extent > block.getBoundingClientRect().height)
    block.style.minHeight = `${extent}px`;
  block.style.zoom = String(Math.min(1, (available - 0.5) / extent));
  block.dataset.docxScaled = "true";
}

/** Scale only an over-height floating drawing, never its anchoring text. */
function fitOversizedDrawing(block: HTMLElement, limit: number): boolean {
  let changed = false;
  for (const drawing of block.querySelectorAll<HTMLElement>(
    "[data-docx-floating-anchor]",
  )) {
    const bounds = drawing.getBoundingClientRect();
    const extent = bottom(drawing) - bounds.top;
    if (extent <= 0 || bounds.bottom <= limit + 0.5) continue;
    const available = limit - bounds.top;
    if (available <= 0) continue;
    drawing.style.zoom = String(Math.min(1, (available - 0.5) / extent));
    drawing.dataset.docxScaled = "true";
    changed = true;
  }
  return changed && fits(block, limit);
}

function cloneShell<T extends HTMLElement>(element: T): T {
  const copy = element.cloneNode(false) as T;
  copy.removeAttribute("id");
  return copy;
}

function repeat(element: HTMLElement): HTMLElement {
  const copy = element.cloneNode(true) as HTMLElement;
  copy.dataset.docxRepeated = "true";
  copy.removeAttribute("id");
  copy.querySelectorAll("[id]").forEach((child) => child.removeAttribute("id"));
  return copy;
}

function textNodes(element: HTMLElement): Text[] {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      return node.parentElement?.closest(".docx-comment-popover")
        ? NodeFilter.FILTER_REJECT
        : NodeFilter.FILTER_ACCEPT;
    },
  });
  const nodes: Text[] = [];
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}

/** Split at the last complete rendered line, preserving inline markup and markers. */
function splitParagraph(
  paragraph: HTMLElement,
  limit: number,
): HTMLElement | null {
  const nodes = textNodes(paragraph);
  let boundary: { node: Text; offset: number } | undefined;
  const range = document.createRange();
  for (const node of nodes) {
    if (!node.length) continue;
    range.selectNodeContents(node);
    const rects = [...range.getClientRects()].filter((r) => r.height > 0);
    if (!rects.length) continue;
    if (rects.every((r) => r.bottom <= limit + 0.5)) {
      boundary = { node, offset: node.length };
      continue;
    }
    let low = 0;
    let high = node.length;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      range.setStart(node, 0);
      range.setEnd(node, middle);
      if ([...range.getClientRects()].every((r) => r.bottom <= limit + 0.5))
        low = middle;
      else high = middle - 1;
    }
    if (low > 0) {
      // Prefer a word boundary. Avoid splitting a surrogate pair for long words.
      const prefix = node.data.slice(0, low);
      const wordEnd = prefix.search(/\s+\S*$/u);
      if (low < node.length && wordEnd > 0) low = wordEnd + 1;
      if (low < node.length && /[\uDC00-\uDFFF]/u.test(node.data[low])) low--;
      if (low) boundary = { node, offset: low };
    }
    break;
  }
  if (!boundary) return null;
  const index = nodes.indexOf(boundary.node);
  if (index === nodes.length - 1 && boundary.offset === boundary.node.length)
    return null;
  range.setStart(boundary.node, boundary.offset);
  range.setEnd(paragraph, paragraph.childNodes.length);
  const continuation = cloneShell(paragraph);
  continuation.append(range.extractContents());
  continuation.dataset.docxContinuation = "true";
  continuation.style.marginTop = "0";
  continuation.style.textIndent = "0";
  continuation.style.breakBefore = "auto";
  paragraph.style.marginBottom = "0";
  return continuation;
}

/** Rowspan-connected rows must move together, including hidden vMerge cells. */
function rowGroups(table: HTMLTableElement): HTMLTableRowElement[][] {
  const rows = [...table.rows];
  const groups: HTMLTableRowElement[][] = [];
  for (let start = 0; start < rows.length;) {
    let end = start + 1;
    for (let i = start; i < end; i++) {
      for (const cell of rows[i].cells) {
        if (getComputedStyle(cell).display !== "none") {
          end = Math.min(
            rows.length,
            Math.max(end, i + (cell.rowSpan || rows.length - i)),
          );
        }
      }
    }
    groups.push(rows.slice(start, end));
    start = end;
  }
  return groups;
}

function tableShell(table: HTMLTableElement): HTMLTableElement {
  const copy = cloneShell(table);
  for (const child of table.children) {
    if (child.tagName === "COLGROUP") copy.append(child.cloneNode(true));
  }
  return copy;
}

function splitTable(
  table: HTMLTableElement,
  limit: number,
  emptyPage: boolean,
): HTMLTableElement | null {
  const groups = rowGroups(table);
  let cut = groups.findIndex((group) => !fits(group[group.length - 1], limit));
  // Isolate an over-height merged group so only that group needs scaling.
  if (cut === 0 && emptyPage && groups.length > 1) cut = 1;
  if (cut <= 0) return null;
  const headers = groups
    .filter((group) =>
      group.every((row) => row.classList.contains("docx-repeat-header")),
    )
    .flat();
  if (
    groups
      .slice(0, cut)
      .flat()
      .every((row) => headers.includes(row))
  ) {
    if (!emptyPage) return null;
    // Keep the header with the oversized first body group. Only that fragment
    // needs scaling; subsequent rows continue at their original text size.
    cut++;
    if (cut >= groups.length) return null;
  }
  const continuation = tableShell(table);
  for (const row of headers) continuation.append(repeat(row));
  for (const group of groups.slice(cut))
    for (const row of group) continuation.append(row);
  return continuation;
}

function splitBlock(
  block: HTMLElement,
  limit: number,
  emptyPage: boolean,
): HTMLElement | null {
  if (block.tagName === "P") return splitParagraph(block, limit);
  if (block instanceof HTMLTableElement)
    return splitTable(block, limit, emptyPage);
  // Lists and other block containers, including footnotes/endnotes.
  if (["OL", "UL", "DIV"].includes(block.tagName)) {
    const children = [...block.children] as HTMLElement[];
    const cut = children.findIndex((child) => !fits(child, limit));
    if (cut > 0) {
      const continuation = cloneShell(block);
      for (const child of children.slice(cut)) continuation.append(child);
      if (continuation instanceof HTMLOListElement)
        continuation.start = (block as HTMLOListElement).start + cut;
      return continuation;
    }
  }
  return null;
}

/** Rebuild ranges because moving/splitting DOM nodes invalidates the library's ranges. */
export function restoreCommentHighlights(container: HTMLElement): void {
  const highlights = CSS.highlights?.get("docx-comments");
  if (!highlights) return;
  highlights.clear();
  const walker = document.createTreeWalker(
    container,
    NodeFilter.SHOW_COMMENT | NodeFilter.SHOW_TEXT,
  );
  const active = new Map<string, Range[]>();
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (
      node.parentElement?.closest("[data-docx-repeated], .docx-comment-popover")
    )
      continue;
    if (node.nodeType === Node.TEXT_NODE) {
      for (const ranges of active.values()) {
        const range = document.createRange();
        range.selectNodeContents(node);
        ranges.push(range);
      }
      continue;
    }
    const marker = /^(start|end) of comment #(.+)$/.exec(
      node.textContent ?? "",
    );
    if (!marker) continue;
    if (marker[1] === "start") active.set(marker[2], []);
    else {
      for (const range of active.get(marker[2]) ?? []) highlights.add(range);
      active.delete(marker[2]);
    }
  }
}

/** Run once after fonts/images settle, before viewport scaling. */
export async function paginateDocx(
  container: HTMLElement,
  isCurrent: () => boolean,
  options: { restoreHighlights?: boolean } = {},
): Promise<number> {
  // Freeze auto-layout column widths before moving rows into partial tables.
  for (const table of container.querySelectorAll<HTMLTableElement>(
    "article table",
  )) {
    const width = table.getBoundingClientRect().width;
    const columns = [
      ...table.querySelectorAll<HTMLElement>(":scope > colgroup > col"),
    ];
    const widths = columns.map((col) => col.getBoundingClientRect().width);
    if (width > 0 && widths.every((w) => w > 0)) {
      table.style.width = `${width}px`;
      table.style.tableLayout = "fixed";
      columns.forEach((col, i) => {
        col.style.width = `${widths[i]}px`;
      });
    }
  }
  const originals = [
    ...container.querySelectorAll<HTMLElement>(".docx-wrapper > section.docx"),
  ];
  let processed = 0;
  for (const source of originals) {
    if (!isCurrent()) return 0;
    const style = getComputedStyle(source);
    const height = px(style.minHeight);
    const paddingBottom = px(style.paddingBottom);
    const header = source.querySelector<HTMLElement>(":scope > header");
    const footer = source.querySelector<HTMLElement>(":scope > footer");
    const decorations = [
      ...source.querySelectorAll<HTMLElement>(
        ":scope > [data-docx-page-decoration]",
      ),
    ];
    const footerStyle = footer && getComputedStyle(footer);
    // Incremental layout clones already-paginated sections, where our CSS
    // resolves margin-top:auto to the remaining free page space. Initial
    // sections can carry a real top margin, which still belongs in the limit.
    const repeatedAutoFooterMargin =
      source.dataset.docxPaginated === "true" &&
      footer?.style.marginTop === "";
    const footerSpace =
      footer && footerStyle
        ? Math.max(
            0,
            footer.offsetHeight +
              (repeatedAutoFooterMargin ? 0 : px(footerStyle.marginTop)) +
              px(footerStyle.marginBottom),
          )
        : 0;
    const streams = [...source.children]
      .filter(
        (child) =>
          child !== header && child !== footer && !decorations.includes(child as HTMLElement),
      )
      .map((child) => {
        const article =
          child.tagName === "ARTICLE"
            ? (child as HTMLElement)
            : document.createElement("article");
        const blocks =
          child.tagName === "ARTICLE" ? [...child.childNodes] : [child];
        return { template: cloneShell(article), blocks };
      });
    source.style.height = `${height}px`;
    source.dataset.docxPaginated = "true";
    for (const child of [...source.children])
      if (
        child !== header &&
        child !== footer &&
        !decorations.includes(child as HTMLElement)
      )
        child.remove();
    let page = source;
    let hasContent = false;
    let article: HTMLElement;
    let articleTemplate: HTMLElement;
    const limit = () =>
      page.getBoundingClientRect().top + height - paddingBottom - footerSpace;
    const addArticle = () => {
      article = cloneShell(articleTemplate);
      page.insertBefore(article, page.querySelector(":scope > footer"));
    };
    const nextPage = () => {
      const next = cloneShell(source);
      if (header) next.append(repeat(header));
      if (footer) next.append(repeat(footer));
      page.after(next);
      page = next;
      hasContent = false;
      addArticle();
    };
    for (const stream of streams) {
      articleTemplate = stream.template;
      addArticle();
      const pending = [...stream.blocks];
      while (pending.length) {
        if (++processed % 64 === 0) {
          await new Promise<void>((resolve) =>
            requestAnimationFrame(() => resolve()),
          );
          if (!isCurrent()) return 0;
        }
        const node = pending.shift()!;
        if (!(node instanceof HTMLElement)) {
          article!.append(node);
          continue;
        }
        article!.append(node);
        const nodeStyle = getComputedStyle(node);
        if (nodeStyle.breakBefore === "page" && hasContent) {
          node.remove();
          nextPage();
          article!.append(node);
        }
        // Keep a heading with at least the first line of the following block.
        const following = pending[0];
        if (
          nodeStyle.breakAfter === "avoid" &&
          following instanceof HTMLElement &&
          hasContent
        ) {
          article!.append(following);
          const firstLine =
            following.getBoundingClientRect().top +
            (px(getComputedStyle(following).lineHeight) || 20);
          following.remove();
          if (firstLine > limit()) {
            node.remove();
            nextPage();
            article!.append(node);
          }
        }
        if (fits(node, limit())) {
          hasContent = true;
          continue;
        }
        if (nodeStyle.breakInside === "avoid" && hasContent) {
          node.remove();
          nextPage();
          article!.append(node);
          if (fits(node, limit())) {
            hasContent = true;
            continue;
          }
        }
        let remainder = splitBlock(node, limit(), !hasContent);
        if (!remainder && hasContent) {
          node.remove();
          nextPage();
          article!.append(node);
          if (fits(node, limit())) {
            hasContent = true;
            continue;
          }
          remainder = splitBlock(node, limit(), true);
        }
        if (remainder) {
          if (
            !fits(node, limit()) &&
            !fitOversizedDrawing(node, limit())
          )
            fitOversized(node, limit());
          hasContent = true;
          nextPage();
          pending.unshift(remainder);
        } else {
          // An indivisible image, drawing, or merged-row group may exceed a page.
          // Fit it visibly instead of clipping it or emitting an unbounded page.
          if (!fitOversizedDrawing(node, limit())) fitOversized(node, limit());
          hasContent = true;
        }
      }
    }
  }
  if (options.restoreHighlights !== false) restoreCommentHighlights(container);
  return container.querySelectorAll(".docx-wrapper > section.docx").length;
}
