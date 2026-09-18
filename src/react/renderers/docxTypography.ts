type CssValues = Record<string, string>;

type DocxStyleRule = {
  target: string;
  values: CssValues;
};

type DocxStyle = {
  id: string | null;
  isDefault?: boolean | null;
  target: string | null;
  basedOn?: string | null;
  linked?: string | null;
  styles: DocxStyleRule[];
  paragraphProps?: {
    tabs?: DocxTabStop[];
    lineSpacing?: DocxLineSpacing;
  };
};

type DocxLineSpacing = {
  line?: number | null;
  lineRule?: string | null;
};

type DocxTabStop = {
  position: string;
  leader?: string | null;
  style?: string | null;
};

type DocxNode = {
  type: string;
  text?: string;
  styleName?: string;
  className?: string;
  cssStyle?: CssValues;
  tabs?: DocxTabStop[];
  lineSpacing?: DocxLineSpacing;
  children?: DocxNode[];
};

type NumberingLevel = {
  id: string;
  level: number;
  levelText?: string;
  start?: number;
};

type ParsedNumberingLevel = {
  level: number;
  start?: string;
  restart?: number;
};

type NumberingOverride = {
  level: number;
  start?: number;
  numberingLevel?: ParsedNumberingLevel;
};

type ConcreteNumbering = {
  id: string;
  abstractId?: string;
  overrides?: NumberingOverride[];
};

type AbstractNumbering = {
  id: string;
  levels: ParsedNumberingLevel[];
};

type DocxDocumentModel = {
  documentPart: { body: DocxNode };
  parts?: Array<{ rootElement?: DocxNode }>;
  partsMap?: Map<string, { rootElement?: DocxNode }> | Record<string, { rootElement?: DocxNode }>;
  stylesPart?: { styles: DocxStyle[] };
  settingsPart?: {
    settings?: { autoHyphenation?: boolean | null; defaultTabStop?: string };
  };
  themePart?: {
    theme?: {
      fontScheme?: {
        majorFont?: { latinTypeface?: string };
        minorFont?: { latinTypeface?: string };
      };
    };
  };
  numberingPart?: {
    domNumberings?: NumberingLevel[];
    numberings?: ConcreteNumbering[];
    abstractNumberings?: AbstractNumbering[];
  };
};

type TabLayout = {
  stops: Array<{ positionPt: number; leader: string; alignment: string }>;
};

export type DocxTypographyState = {
  defaultTabStopPt: number;
  tabs: Map<string, TabLayout>;
};

const TAB_CLASS = "docx-owned-tab-";

function withArialFallback(value: string): string {
  if (/\bsans-serif\b/i.test(value)) return value;
  return `${value}, Arial, sans-serif`;
}

function addFontFallback(values: CssValues | undefined): void {
  if (values?.["font-family"])
    values["font-family"] = withArialFallback(values["font-family"]);
}

function resolvedStyleValues(
  style: DocxStyle,
  target: string,
  styles: Map<string, DocxStyle>,
  seen = new Set<string>(),
): CssValues {
  const result: CssValues = {};
  if (style.basedOn && !seen.has(style.basedOn)) {
    seen.add(style.basedOn);
    const base = styles.get(style.basedOn);
    if (base) Object.assign(result, resolvedStyleValues(base, target, styles, seen));
  }
  for (const rule of style.styles)
    if (rule.target === target) Object.assign(result, rule.values);
  return result;
}

function styleTabs(
  styleName: string | undefined,
  styles: Map<string, DocxStyle>,
  seen = new Set<string>(),
): DocxTabStop[] | undefined {
  if (!styleName || seen.has(styleName)) return undefined;
  seen.add(styleName);
  const style = styles.get(styleName);
  return style?.paragraphProps?.tabs ?? styleTabs(style?.basedOn ?? undefined, styles, seen);
}

function styleLineSpacing(
  styleName: string | undefined,
  styles: Map<string, DocxStyle>,
  seen = new Set<string>(),
): DocxLineSpacing | undefined {
  if (!styleName || seen.has(styleName)) return undefined;
  seen.add(styleName);
  const style = styles.get(styleName);
  const inherited = styleLineSpacing(style?.basedOn ?? undefined, styles, seen);
  const own = style?.paragraphProps?.lineSpacing;
  const line = own?.line ?? inherited?.line;
  const lineRule = own?.lineRule ?? inherited?.lineRule;
  return line != null || lineRule != null ? { line, lineRule } : undefined;
}

function restoreAtLeastLineSpacing(
  paragraph: DocxNode,
  styles: Map<string, DocxStyle>,
  defaultStyleName: string | undefined,
): void {
  if (paragraph.className?.split(/\s+/u).includes("docx-section-carrier")) return;

  const defaults = styleLineSpacing(defaultStyleName, styles);
  const named = styleLineSpacing(paragraph.styleName, styles);
  const line = paragraph.lineSpacing?.line ?? named?.line ?? defaults?.line;
  const lineRule =
    paragraph.lineSpacing?.lineRule ?? named?.lineRule ?? defaults?.lineRule;
  if (lineRule !== "atLeast" || line == null) return;

  const points = line / 20;
  if (!Number.isFinite(points) || points < 0) return;

  const minimum = `${points}pt`;
  paragraph.cssStyle ??= {};
  paragraph.cssStyle["line-height"] = minimum;
  paragraph.cssStyle["min-height"] = minimum;

  // CSS gives a block's line-height strut to every inline descendant. Reset
  // those descendants to their natural font line height so the paragraph
  // value acts as a minimum, matching Word's atLeast rule. Larger runs can
  // still expand the line box above the requested minimum.
  function restoreNaturalHeight(node: DocxNode): void {
    if (node.type !== "text" && node.type !== "break" && node.type !== "tab") {
      node.cssStyle ??= {};
      node.cssStyle["line-height"] = "normal";
    }
    node.children?.forEach(restoreNaturalHeight);
  }
  paragraph.children?.forEach(restoreNaturalHeight);
}

function documentRoots(document: DocxDocumentModel): DocxNode[] {
  const roots = [document.documentPart.body];
  for (const part of document.parts ?? [])
    if (part.rootElement) roots.push(part.rootElement);
  const mappedParts = document.partsMap;
  if (mappedParts instanceof Map) {
    for (const part of mappedParts.values())
      if (part.rootElement) roots.push(part.rootElement);
  } else if (mappedParts) {
    for (const part of Object.values(mappedParts))
      if (part.rootElement) roots.push(part.rootElement);
  }
  return [...new Set(roots)];
}

function parsePoints(value: string | undefined, fallback = 0): number {
  const parsed = Number.parseFloat(value ?? "");
  return Number.isFinite(parsed) ? parsed : fallback;
}

/**
 * Repairs docx-preview's Word-style cascade and marks tabs for deterministic
 * layout. Call this after parseAsync and before renderDocument.
 */
export function normalizeDocxTypography(
  document: DocxDocumentModel,
): DocxTypographyState {
  const styles = document.stylesPart?.styles ?? [];
  const stylesById = new Map(
    styles.flatMap((style) => (style.id ? [[style.id, style] as const] : [])),
  );

  // docx-preview applies a linked character style to the paragraph itself.
  // Word uses that linked style only when a run explicitly selects it.
  for (const style of styles)
    if (style.target === "p") style.linked = null;

  // The library emits docDefaults directly on every span, but emits the
  // default paragraph style on the parent paragraph. Collapse Word's cascade
  // so Normal's run formatting wins on otherwise unstyled runs.
  const defaultParagraph = styles.find(
    (style) => style.target === "p" && style.isDefault,
  );
  const documentDefaults = styles.find(
    (style) => style.id == null && style.target == null,
  );
  if (defaultParagraph && documentDefaults) {
    const defaultRun = documentDefaults.styles.find(
      (rule) => rule.target === "span",
    );
    if (defaultRun)
      Object.assign(
        defaultRun.values,
        resolvedStyleValues(defaultParagraph, "span", stylesById),
      );
  }

  const defaultParagraphValues = defaultParagraph
    ? resolvedStyleValues(defaultParagraph, "p", stylesById)
    : undefined;

  const scheme = document.themePart?.theme?.fontScheme;
  for (const font of [scheme?.majorFont, scheme?.minorFont]) {
    if (font?.latinTypeface)
      font.latinTypeface = withArialFallback(font.latinTypeface);
  }
  for (const style of styles)
    for (const rule of style.styles) addFontFallback(rule.values);
  for (const level of document.numberingPart?.domNumberings ?? [])
    addFontFallback((level as NumberingLevel & { rStyle?: CssValues }).rStyle);

  const state: DocxTypographyState = {
    defaultTabStopPt: Math.max(
      1,
      parsePoints(document.settingsPart?.settings?.defaultTabStop, 36),
    ),
    tabs: new Map(),
  };
  let tabId = 0;

  function markTabs(paragraph: DocxNode): void {
    const stops = paragraph.tabs ?? styleTabs(paragraph.styleName, stylesById) ?? [];
    function visit(node: DocxNode): void {
      addFontFallback(node.cssStyle);
      if (!node.children) return;
      node.children = node.children.map((child) => {
        if (child.type !== "tab") {
          visit(child);
          return child;
        }
        const id = String(tabId++);
        state.tabs.set(id, {
          stops: stops
            .map((stop) => ({
              positionPt: parsePoints(stop.position),
              leader: stop.leader ?? "none",
              alignment: stop.style ?? "left",
            }))
            .filter((stop) => stop.positionPt >= 0)
            .sort((a, b) => a.positionPt - b.positionPt),
        });
        return {
          type: "run",
          className: `docx-tab-stop ${TAB_CLASS}${id}`,
          children: [{ type: "text", text: "\u00a0" }],
        };
      });
    }
    visit(paragraph);
  }

  function walk(node: DocxNode): void {
    addFontFallback(node.cssStyle);
    if (node.type === "paragraph") {
      // A table style selector is more specific than a paragraph style
      // selector in CSS, while Word gives paragraph formatting precedence.
      // Put the resolved paragraph cascade inline so the browser follows
      // Word's ordering: Normal, named style, direct pPr. Document defaults
      // remain in the library stylesheet below table-style precedence.
      const paragraphStyle = node.styleName
        ? stylesById.get(node.styleName)
        : undefined;
      node.cssStyle = {
        ...defaultParagraphValues,
        ...(paragraphStyle
          ? resolvedStyleValues(paragraphStyle, "p", stylesById)
          : undefined),
        ...node.cssStyle,
      };
      restoreAtLeastLineSpacing(
        node,
        stylesById,
        defaultParagraph?.id ?? undefined,
      );
      markTabs(node);
    }
    node.children?.forEach(walk);
  }
  documentRoots(document).forEach(walk);
  return state;
}

function concreteRestart(
  document: DocxDocumentModel,
  numberingId: string,
  level: number,
): number | undefined {
  const part = document.numberingPart;
  const concrete = part?.numberings?.find((item) => item.id === numberingId);
  const override = concrete?.overrides?.find((item) => item.level === level);
  if (override?.numberingLevel?.restart !== undefined)
    return override.numberingLevel.restart;
  const abstract = part?.abstractNumberings?.find(
    (item) => item.id === concrete?.abstractId,
  );
  return abstract?.levels.find((item) => item.level === level)?.restart;
}

function concreteStart(
  document: DocxDocumentModel,
  numberingId: string,
  level: NumberingLevel,
): number {
  const concrete = document.numberingPart?.numberings?.find(
    (item) => item.id === numberingId,
  );
  const override = concrete?.overrides?.find((item) => item.level === level.level);
  return (
    override?.start ??
    (override?.numberingLevel?.start
      ? Number.parseInt(override.numberingLevel.start, 10)
      : undefined) ??
    level.start ??
    1
  );
}

/**
 * Adds the deeper-level resets that docx-preview omits when a list skips a
 * level. Append this CSS after the library's generated numbering styles.
 */
export function createDocxNumberingFixStyles(
  document: DocxDocumentModel,
): string {
  const levelsByNumbering = new Map<string, NumberingLevel[]>();
  for (const level of document.numberingPart?.domNumberings ?? []) {
    if (!/^[\w-]+$/u.test(level.id)) continue;
    const levels = levelsByNumbering.get(level.id) ?? [];
    levels.push(level);
    levelsByNumbering.set(level.id, levels);
  }

  const rules: string[] = [];
  for (const [id, levels] of levelsByNumbering) {
    levels.sort((a, b) => a.level - b.level);
    for (const trigger of levels) {
      const resets: string[] = [];
      for (const target of levels) {
        if (target.level <= trigger.level) continue;
        const restart = concreteRestart(document, id, target.level);
        if (restart === 0) continue;
        // OOXML's explicit value is one-based. When omitted, Word resets a
        // deeper counter after any higher level, including skipped levels.
        if (restart !== undefined && trigger.level !== restart - 1) continue;
        resets.push(
          `docx-num-${id}-${target.level} ${concreteStart(document, id, target) - 1}`,
        );
      }
      const librarySetsImmediateCounter = levels.some(
        (target) =>
          target.level === trigger.level + 1 && Boolean(target.levelText),
      );
      if (resets.length || librarySetsImmediateCounter) {
        rules.push(
          `.docx-render-target p.docx-num-${id}-${trigger.level}:not([data-docx-continuation]) { counter-set: ${resets.join(" ") || "none"}; }`,
        );
      }
      if (librarySetsImmediateCounter) {
        rules.push(
          `.docx-render-target p.docx-num-${id}-${trigger.level}[data-docx-continuation] { counter-set: none; }`,
        );
      }
    }
  }
  return rules.join("\n");
}

/** Match Word's document-level auto-hyphenation setting. */
export function createDocxHyphenationStyles(
  document: DocxDocumentModel,
): string {
  const hyphens =
    document.settingsPart?.settings?.autoHyphenation === true
      ? "auto"
      : "manual";
  return `.docx-render-target section.docx { hyphens: ${hyphens}; }`;
}

function tabId(element: Element): string | undefined {
  return [...element.classList]
    .find((name) => name.startsWith(TAB_CLASS))
    ?.slice(TAB_CLASS.length);
}

function pixelsPerPoint(container: HTMLElement): number {
  const ruler = document.createElement("span");
  Object.assign(ruler.style, {
    display: "block",
    height: "0",
    position: "absolute",
    visibility: "hidden",
    width: "100pt",
  });
  container.append(ruler);
  const scale = ruler.getBoundingClientRect().width / 100;
  ruler.remove();
  return scale || 96 / 72;
}

function followingWidth(
  tab: HTMLElement,
  nextTab: HTMLElement | undefined,
  paragraph: HTMLParagraphElement,
): number {
  const range = document.createRange();
  range.setStartAfter(tab);
  if (nextTab) range.setEndBefore(nextTab);
  else range.setEnd(paragraph, paragraph.childNodes.length);
  return range.getBoundingClientRect().width;
}

function setLeader(tab: HTMLElement, leader: string): void {
  tab.style.backgroundImage = "";
  tab.style.backgroundPosition = "";
  tab.style.backgroundRepeat = "";
  tab.style.backgroundSize = "";
  if (leader === "dot" || leader === "middleDot") {
    tab.style.backgroundImage =
      "radial-gradient(circle, currentColor 0.75px, transparent 0.85px)";
    tab.style.backgroundPosition = "left 75%";
    tab.style.backgroundRepeat = "repeat-x";
    tab.style.backgroundSize = "4px 2px";
  } else if (["hyphen", "heavy", "underscore"].includes(leader)) {
    tab.style.backgroundImage =
      "linear-gradient(to right, currentColor 60%, transparent 60%)";
    tab.style.backgroundPosition = "left 78%";
    tab.style.backgroundRepeat = "repeat-x";
    tab.style.backgroundSize = "6px 1px";
  }
}

function nextDefaultStop(
  current: number,
  interval: number,
  stops: TabLayout["stops"],
): number {
  let position = (Math.floor(current / interval) + 1) * interval;
  while (
    stops.some(
      (stop) =>
        stop.alignment === "clear" && Math.abs(stop.positionPt - position) < 0.1,
    )
  )
    position += interval;
  return position;
}

/** Lay out marked OOXML tabs after fonts settle. Safe to rerun after pagination. */
export function layoutDocxTabs(
  container: HTMLElement,
  state: DocxTypographyState,
): void {
  const tabs = [
    ...container.querySelectorAll<HTMLElement>(`.docx-tab-stop[class*="${TAB_CLASS}"]`),
  ];
  for (const tab of tabs) {
    tab.style.display = "inline-block";
    tab.style.width = "0";
    tab.style.whiteSpace = "nowrap";
    setLeader(tab, "none");
  }
  const pointScale = pixelsPerPoint(container);
  const tabsByParagraph = new Map<HTMLParagraphElement, HTMLElement[]>();
  for (const tab of tabs) {
    const paragraph = tab.closest("p");
    if (!paragraph) continue;
    const paragraphTabs = tabsByParagraph.get(paragraph) ?? [];
    paragraphTabs.push(tab);
    tabsByParagraph.set(paragraph, paragraphTabs);
  }

  for (const [paragraph, paragraphTabs] of tabsByParagraph) {
    const paragraphIndentPixels = Number.parseFloat(
      getComputedStyle(paragraph).marginInlineStart,
    );
    const paragraphIndent = Number.isFinite(paragraphIndentPixels)
      ? paragraphIndentPixels / pointScale
      : 0;
    for (let index = 0; index < paragraphTabs.length; index++) {
      const tab = paragraphTabs[index];
      const definition = state.tabs.get(tabId(tab) ?? "");
      if (!definition) continue;
      const current =
        paragraphIndent +
        (tab.getBoundingClientRect().left -
          paragraph.getBoundingClientRect().left) /
          pointScale;
      const explicit = definition.stops.find(
        (stop) => stop.alignment !== "clear" && stop.positionPt > current + 0.1,
      );
      const stop =
        explicit ??
        ({
          positionPt: nextDefaultStop(
            current,
            state.defaultTabStopPt,
            definition.stops,
          ),
          leader: "none",
          alignment: "left",
        } satisfies TabLayout["stops"][number]);
      const tail =
        followingWidth(tab, paragraphTabs[index + 1], paragraph) / pointScale;
      const alignedTail =
        stop.alignment === "right"
          ? tail
          : stop.alignment === "center"
            ? tail / 2
            : 0;
      tab.style.width = `${Math.max(0, stop.positionPt - current - alignedTail)}pt`;
      setLeader(tab, stop.leader);
    }
  }
}
