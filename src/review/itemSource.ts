import type {
  DocxManifest,
  ItemDocxSource,
  LiveDocxAnnotation,
  LiveDocxListener,
  LiveDocxParagraph,
  LiveDocxSnapshot,
  ProjectedEntity,
  ReviewItem,
} from "./model";
import { readDocx } from "./readDocx";
import { resolveAnchors } from "./resolveAnchors";

export type { ItemDocxSource } from "./model";

function now(): number {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

function cloneItems(items: readonly ReviewItem[]): readonly ReviewItem[] {
  return items.map((item) => {
    const value = item as unknown;
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return value as ReviewItem;
    }
    const candidate = value as Record<string, unknown>;
    const anchor = candidate.anchor;
    return {
      ...candidate,
      ...(anchor && typeof anchor === "object" && !Array.isArray(anchor)
        ? { anchor: { ...(anchor as Record<string, unknown>) } }
        : {}),
    } as ReviewItem;
  });
}

function itemValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (
    !left ||
    !right ||
    typeof left !== "object" ||
    typeof right !== "object" ||
    Array.isArray(left) ||
    Array.isArray(right)
  ) {
    return false;
  }
  const a = left as Partial<ReviewItem>;
  const b = right as Partial<ReviewItem>;
  const leftAnchor = a.anchor;
  const rightAnchor = b.anchor;
  const anchorsEqual =
    Object.is(leftAnchor, rightAnchor) ||
    (!!leftAnchor &&
      !!rightAnchor &&
      typeof leftAnchor === "object" &&
      typeof rightAnchor === "object" &&
      leftAnchor.quote === rightAnchor.quote &&
      leftAnchor.prefix === rightAnchor.prefix &&
      leftAnchor.suffix === rightAnchor.suffix &&
      leftAnchor.paragraphId === rightAnchor.paragraphId);
  return (
    a.id === b.id &&
    a.kind === b.kind &&
    a.body === b.body &&
    a.replacement === b.replacement &&
    Object.is(a.metadata, b.metadata) &&
    anchorsEqual
  );
}

function itemsEqual(
  left: readonly ReviewItem[],
  right: readonly ReviewItem[],
): boolean {
  return (
    left.length === right.length &&
    left.every((item, index) => itemValueEqual(item, right[index]))
  );
}

function annotationsEqual(
  left: readonly LiveDocxAnnotation[],
  right: readonly LiveDocxAnnotation[],
): boolean {
  return (
    left.length === right.length &&
    left.every((value, index) => {
      const other = right[index];
      return (
        value.entityId === other.entityId &&
        value.kind === other.kind &&
        value.start === other.start &&
        value.end === other.end &&
        value.first === other.first &&
        value.last === other.last
      );
    })
  );
}

function paragraphsEqual(
  left: LiveDocxParagraph,
  right: LiveDocxParagraph,
): boolean {
  return (
    left.texts.length === right.texts.length &&
    left.texts.every((value, index) => {
      const other = right.texts[index];
      return (
        value.id === other.id &&
        value.value === other.value &&
        annotationsEqual(value.annotations, other.annotations)
      );
    })
  );
}

function entityEqual(left: ProjectedEntity, right: ProjectedEntity): boolean {
  return (
    left.id === right.id &&
    left.kind === right.kind &&
    left.body === right.body &&
    left.quote === right.quote &&
    left.replacement === right.replacement &&
    Object.is(left.metadata, right.metadata) &&
    left.status === right.status &&
    left.reason === right.reason
  );
}

function changedEntityIds(
  previous: readonly ProjectedEntity[],
  next: readonly ProjectedEntity[],
): Set<string> {
  const previousById = new Map<string, ProjectedEntity[]>();
  const nextById = new Map<string, ProjectedEntity[]>();
  for (const entity of previous) {
    previousById.set(entity.id, [...(previousById.get(entity.id) ?? []), entity]);
  }
  for (const entity of next) {
    nextById.set(entity.id, [...(nextById.get(entity.id) ?? []), entity]);
  }
  const changed = new Set<string>();
  for (const id of new Set([...previousById.keys(), ...nextById.keys()])) {
    const before = previousById.get(id) ?? [];
    const after = nextById.get(id) ?? [];
    if (
      before.length !== after.length ||
      before.some((entity, index) => !entityEqual(entity, after[index]))
    ) {
      changed.add(id);
    }
  }
  return changed;
}

function addEntityParagraphs(
  result: Set<string>,
  paragraphs: ReadonlyMap<string, LiveDocxParagraph>,
  entityIds: ReadonlySet<string>,
): void {
  if (entityIds.size === 0) return;
  for (const paragraph of paragraphs.values()) {
    if (
      paragraph.texts.some((text) =>
        text.annotations.some((annotation) => entityIds.has(annotation.entityId)),
      )
    ) {
      result.add(paragraph.id);
    }
  }
}

class ExternalItemDocxSource implements ItemDocxSource {
  readonly file: File;
  readonly manifest: DocxManifest;

  private readonly texts: Readonly<Record<string, string>>;
  private readonly listeners = new Set<LiveDocxListener>();
  private items: readonly ReviewItem[];
  private pendingItems: readonly ReviewItem[] | null = null;
  private revision = 0;
  private projectionQueued = false;
  private disposed = false;
  private snapshot: LiveDocxSnapshot;

  constructor(
    file: File,
    manifest: ExternalItemDocxSource["manifest"],
    texts: Readonly<Record<string, string>>,
    items: readonly ReviewItem[],
  ) {
    this.file = file;
    this.manifest = manifest;
    this.texts = texts;
    this.items = cloneItems(items);
    const began = now();
    const projection = resolveAnchors(this.manifest, this.texts, this.items);
    this.snapshot = {
      revision: 0,
      projectionMs: Math.round(now() - began),
      ...projection,
    };
  }

  getSnapshot = (): LiveDocxSnapshot => this.snapshot;

  subscribe = (listener: LiveDocxListener): (() => void) => {
    if (this.disposed) return () => undefined;
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  setItems = (items: readonly ReviewItem[]): void => {
    if (this.disposed) return;
    const next = cloneItems(items);
    const compared = this.pendingItems ?? this.items;
    if (itemsEqual(compared, next)) return;
    this.pendingItems = next;
    this.revision += 1;
    if (!this.projectionQueued) {
      this.projectionQueued = true;
      queueMicrotask(this.flushProjection);
    }
  };

  dispose = (): void => {
    if (this.disposed) return;
    this.disposed = true;
    this.pendingItems = null;
    this.listeners.clear();
  };

  private readonly flushProjection = (): void => {
    this.projectionQueued = false;
    if (this.disposed || !this.pendingItems) return;
    const items = this.pendingItems;
    this.pendingItems = null;
    const revision = this.revision;
    if (itemsEqual(this.items, items)) return;
    const began = now();
    let projection;
    try {
      projection = resolveAnchors(this.manifest, this.texts, items);
    } catch (error) {
      console.error("Review item projection failed", error);
      return;
    }

    const previous = this.snapshot;
    const paragraphs = new Map<string, LiveDocxParagraph>();
    const changedParagraphIds = new Set<string>();
    for (const [id, paragraph] of projection.paragraphs) {
      const existing = previous.paragraphs.get(id);
      if (existing && paragraphsEqual(existing, paragraph)) {
        paragraphs.set(id, existing);
      } else {
        paragraphs.set(id, paragraph);
        changedParagraphIds.add(id);
      }
    }
    for (const id of previous.paragraphs.keys()) {
      if (!paragraphs.has(id)) changedParagraphIds.add(id);
    }
    const entityChanges = changedEntityIds(previous.entities, projection.entities);
    addEntityParagraphs(changedParagraphIds, previous.paragraphs, entityChanges);
    addEntityParagraphs(changedParagraphIds, paragraphs, entityChanges);

    this.items = items;
    this.snapshot = {
      revision,
      projectionMs: Math.round(now() - began),
      paragraphs,
      entities: projection.entities,
    };
    const change = { snapshot: this.snapshot, changedParagraphIds };
    for (const listener of this.listeners) {
      try {
        listener(change);
      } catch (error) {
        console.error("Review item source listener failed", error);
      }
    }

  };
}

export async function createItemSource(
  file: File,
  items: readonly ReviewItem[],
): Promise<ItemDocxSource> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const { manifest, texts } = await readDocx(bytes);
  return new ExternalItemDocxSource(file, manifest, texts, items);
}
