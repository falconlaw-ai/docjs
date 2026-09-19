import type {
  DocxManifest,
  DocxParagraph,
  LiveDocxAnnotation,
  LiveDocxParagraph,
  ProjectedEntity,
  ReviewItem,
} from "./model";

const MAX_ITEMS = 5_000;
const MAX_TEXT_LENGTH = 1_000_000;

type ParagraphPart =
  | { kind: "text"; textId: string; start: number; end: number }
  | { kind: "literal"; value: "\t" | "\n"; start: number; end: number };

type SearchableParagraph = {
  descriptor: DocxParagraph;
  value: string;
  parts: ParagraphPart[];
};

type ResolvedSegment = {
  textId: string;
  start: number;
  end: number;
};

type ResolvedItem = {
  item?: ReviewItem;
  entity: ProjectedEntity;
  paragraphId?: string;
  start?: number;
  end?: number;
  segments?: ResolvedSegment[];
};

export type ReviewProjection = {
  paragraphs: ReadonlyMap<string, LiveDocxParagraph>;
  entities: readonly ProjectedEntity[];
};

function itemId(value: unknown, index: number): string {
  if (value && typeof value === "object") {
    const id = (value as { id?: unknown }).id;
    if (typeof id === "string" && id) return id;
  }
  return `invalid:${index}`;
}

function projected(value: unknown, id: string): ProjectedEntity {
  const item =
    value && typeof value === "object" && !Array.isArray(value)
      ? (value as Partial<ReviewItem>)
      : {};
  const kind = item.kind === "redline" ? "redline" : "comment";
  const quote =
    item.anchor && typeof item.anchor.quote === "string"
      ? item.anchor.quote
      : "";
  return {
    id,
    kind,
    ...(typeof item.body === "string" ? { body: item.body } : {}),
    quote,
    ...(typeof item.replacement === "string"
      ? { replacement: item.replacement }
      : {}),
    ...(item.metadata === undefined ? {} : { metadata: item.metadata }),
    status: "open",
  };
}

function invalidate(entity: ProjectedEntity, reason: string): void {
  entity.status = "invalid";
  entity.reason = reason;
}

function validateItem(value: unknown, id: string): string | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return "Review item is not an object";
  }
  const item = value as Partial<ReviewItem>;
  if (typeof item.id !== "string" || !item.id || item.id.length > 256) {
    return "Review item ID is invalid";
  }
  if (item.id !== id) return "Review item ID is invalid";
  if (item.kind !== "comment" && item.kind !== "redline") {
    return "Review item kind is invalid";
  }
  if (!item.anchor || typeof item.anchor !== "object") {
    return "Review item anchor is missing";
  }
  if (
    typeof item.anchor.quote !== "string" ||
    !item.anchor.quote ||
    item.anchor.quote.length > MAX_TEXT_LENGTH
  ) {
    return "Review item quote is invalid";
  }
  for (const name of ["prefix", "suffix"] as const) {
    const context = item.anchor[name];
    if (
      context !== undefined &&
      (typeof context !== "string" || context.length > MAX_TEXT_LENGTH)
    ) {
      return `Review item ${name} is invalid`;
    }
  }
  if (
    item.anchor.paragraphId !== undefined &&
    (typeof item.anchor.paragraphId !== "string" ||
      !item.anchor.paragraphId ||
      item.anchor.paragraphId.length > 256)
  ) {
    return "Review item paragraph ID is invalid";
  }
  if (item.body !== undefined && typeof item.body !== "string") {
    return "Review item body is invalid";
  }
  if (item.kind === "redline") {
    if (typeof item.replacement !== "string") {
      return "Redline replacement is missing";
    }
    if (item.replacement.length > MAX_TEXT_LENGTH) {
      return "Redline replacement is too large";
    }
    if (/[\r\n\t]/.test(item.replacement)) {
      return "Redline replacements containing tabs or line breaks are unsupported";
    }
  } else if (
    item.replacement !== undefined &&
    typeof item.replacement !== "string"
  ) {
    return "Review item replacement is invalid";
  }
  return undefined;
}

function searchableParagraphs(
  manifest: DocxManifest,
  texts: Readonly<Record<string, string>>,
): SearchableParagraph[] {
  return manifest.paragraphs.map((descriptor) => {
    const parts: ParagraphPart[] = [];
    let value = "";
    for (const segment of descriptor.segments) {
      const start = value.length;
      if ("textId" in segment) {
        const text = texts[segment.textId];
        if (typeof text !== "string") {
          throw new Error(`DOCX baseline text ${segment.textId} is missing`);
        }
        value += text;
        parts.push({ kind: "text", textId: segment.textId, start, end: value.length });
      } else {
        value += segment.literal;
        parts.push({
          kind: "literal",
          value: segment.literal,
          start,
          end: value.length,
        });
      }
    }
    return { descriptor, value, parts };
  });
}

function occurrences(value: string, quote: string): number[] {
  const result: number[] = [];
  for (let from = 0; from <= value.length - quote.length; ) {
    const index = value.indexOf(quote, from);
    if (index < 0) break;
    result.push(index);
    from = index + 1;
  }
  return result;
}

function contextMatches(
  paragraph: string,
  start: number,
  end: number,
  prefix: string | undefined,
  suffix: string | undefined,
): boolean {
  return (
    (prefix === undefined ||
      (start >= prefix.length &&
        paragraph.slice(start - prefix.length, start) === prefix)) &&
    (suffix === undefined || paragraph.slice(end, end + suffix.length) === suffix)
  );
}

function crossesParagraphBoundary(
  item: ReviewItem,
  paragraphs: readonly SearchableParagraph[],
): boolean {
  const hint = item.anchor.paragraphId;
  for (let index = 0; index < paragraphs.length - 1; index += 1) {
    const left = paragraphs[index];
    const right = paragraphs[index + 1];
    if (
      hint &&
      left.descriptor.id !== hint &&
      left.descriptor.nativeId !== hint
    ) {
      continue;
    }
    const combined = left.value + right.value;
    for (const start of occurrences(combined, item.anchor.quote)) {
      const end = start + item.anchor.quote.length;
      if (
        start < left.value.length &&
        end > left.value.length &&
        contextMatches(
          combined,
          start,
          end,
          item.anchor.prefix,
          item.anchor.suffix,
        )
      ) {
        return true;
      }
    }
  }
  return false;
}

function resolveItem(
  item: ReviewItem,
  entity: ProjectedEntity,
  paragraphs: readonly SearchableParagraph[],
): ResolvedItem {
  const hint = item.anchor.paragraphId;
  const candidates = hint
    ? paragraphs.filter(
        ({ descriptor }) => descriptor.id === hint || descriptor.nativeId === hint,
      )
    : paragraphs;
  if (hint && candidates.length === 0) {
    invalidate(entity, `Paragraph hint ${hint} does not match this document revision`);
    return { item, entity };
  }
  if (hint && candidates.length > 1) {
    invalidate(entity, `Paragraph hint ${hint} matches multiple paragraphs`);
    return { item, entity };
  }

  const matches: Array<{
    paragraph: SearchableParagraph;
    start: number;
    end: number;
  }> = [];
  for (const paragraph of candidates) {
    for (const start of occurrences(paragraph.value, item.anchor.quote)) {
      const end = start + item.anchor.quote.length;
      if (
        contextMatches(
          paragraph.value,
          start,
          end,
          item.anchor.prefix,
          item.anchor.suffix,
        )
      ) {
        matches.push({ paragraph, start, end });
      }
    }
  }
  if (matches.length === 0) {
    if (crossesParagraphBoundary(item, paragraphs)) {
      invalidate(entity, "Cross-paragraph anchors are unsupported");
      return { item, entity };
    }
    invalidate(entity, "Anchor quote and context do not match the source document");
    return { item, entity };
  }
  if (matches.length > 1) {
    invalidate(entity, "Anchor quote and context are ambiguous in the source document");
    return { item, entity };
  }

  const [{ paragraph, start, end }] = matches;
  if (paragraph.descriptor.editable === false) {
    invalidate(
      entity,
      paragraph.descriptor.reason ?? "Anchor belongs to unsupported Word content",
    );
    return { item, entity };
  }
  if (
    paragraph.parts.some(
      (part) => part.kind === "literal" && part.start < end && part.end > start,
    )
  ) {
    invalidate(entity, "Anchors crossing tabs or breaks are unsupported");
    return { item, entity };
  }

  const segments = paragraph.parts.flatMap((part): ResolvedSegment[] => {
    if (part.kind !== "text" || part.start >= end || part.end <= start) return [];
    return [
      {
        textId: part.textId,
        start: Math.max(start, part.start) - part.start,
        end: Math.min(end, part.end) - part.start,
      },
    ];
  });
  if (segments.length === 0) {
    invalidate(entity, "Anchor does not cover supported Word text");
    return { item, entity };
  }
  return {
    item,
    entity,
    paragraphId: paragraph.descriptor.id,
    start,
    end,
    segments,
  };
}

function redlinesOverlap(left: ResolvedItem, right: ResolvedItem): boolean {
  return (
    left.paragraphId !== undefined &&
    left.paragraphId === right.paragraphId &&
    left.start !== undefined &&
    left.end !== undefined &&
    right.start !== undefined &&
    right.end !== undefined &&
    Math.max(left.start, right.start) < Math.min(left.end, right.end)
  );
}

export function resolveAnchors(
  manifest: DocxManifest,
  texts: Readonly<Record<string, string>>,
  items: readonly ReviewItem[],
): ReviewProjection {
  const overLimit = items.length > MAX_ITEMS;
  const paragraphs = searchableParagraphs(manifest, texts);
  const idCounts = new Map<string, number>();
  items.forEach((item, index) => {
    const id = itemId(item, index);
    idCounts.set(id, (idCounts.get(id) ?? 0) + 1);
  });

  const resolved = items.map((value, index): ResolvedItem => {
    const id = itemId(value, index);
    const entity = projected(value, id);
    const reason = validateItem(value, id);
    if (reason) {
      invalidate(entity, reason);
      return { entity };
    }
    const item = value as ReviewItem;
    if (overLimit) {
      invalidate(entity, `Review input contains more than ${MAX_ITEMS} items`);
      return { item, entity };
    }
    if ((idCounts.get(id) ?? 0) > 1) {
      invalidate(entity, `Review item ID ${id} is duplicated`);
      return { item, entity };
    }
    return resolveItem(item, entity, paragraphs);
  });

  const redlines = resolved.filter(
    ({ item, entity, paragraphId }) =>
      item?.kind === "redline" &&
      entity.status === "open" &&
      paragraphId !== undefined,
  );
  for (let left = 0; left < redlines.length; left += 1) {
    for (let right = left + 1; right < redlines.length; right += 1) {
      if (!redlinesOverlap(redlines[left], redlines[right])) continue;
      for (const entry of [redlines[left], redlines[right]]) {
        entry.entity.status = "conflict";
        entry.entity.reason = "Redline overlaps another redline";
      }
    }
  }

  const annotations = new Map<string, LiveDocxAnnotation[]>();
  for (const entry of resolved) {
    if (entry.entity.status !== "open" || !entry.segments) continue;
    entry.segments.forEach((segment, index) => {
      const values = annotations.get(segment.textId) ?? [];
      values.push({
        entityId: entry.entity.id,
        kind: entry.entity.kind,
        start: segment.start,
        end: segment.end,
        first: index === 0,
        last: index === entry.segments!.length - 1,
      });
      annotations.set(segment.textId, values);
    });
  }
  for (const values of annotations.values()) {
    values.sort(
      (left, right) =>
        left.start - right.start ||
        right.end - left.end ||
        left.kind.localeCompare(right.kind) ||
        left.entityId.localeCompare(right.entityId),
    );
  }

  const liveParagraphs = new Map<string, LiveDocxParagraph>();
  for (const paragraph of manifest.paragraphs) {
    liveParagraphs.set(paragraph.id, {
      id: paragraph.id,
      texts: paragraph.segments.flatMap((segment) => {
        if (!("textId" in segment)) return [];
        const value = texts[segment.textId];
        if (typeof value !== "string") {
          throw new Error(`DOCX baseline text ${segment.textId} is missing`);
        }
        return [
          {
            id: segment.textId,
            value,
            annotations: annotations.get(segment.textId) ?? [],
          },
        ];
      }),
    });
  }
  return {
    paragraphs: liveParagraphs,
    entities: resolved.map(({ entity }) => entity),
  };
}
