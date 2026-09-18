import type {
  DocxManifest,
  LiveDocxParagraph,
  LiveDocxSnapshot,
  ProjectedEntity,
} from "../../review/model";
import { paginateDocx, restoreCommentHighlights } from "./paginateDocx";

type ModelNode = {
  type: string;
  id?: string;
  className?: string;
  text?: string;
  children?: ModelNode[];
};

type ModelDocument = {
  documentPart: { body: ModelNode };
};

const RUN_ID_PREFIX = "falcon-live-text-";
const PARAGRAPH_CLASS_PREFIX = "falcon-live-paragraph-";
const FIELD_TEXT_CLASS_PREFIX = "falcon-live-field-text-";
type InsertionFormat = { className: string; cssText: string };
export type DocxRevisionMode = "review" | "final";

function paragraphClass(id: string): string {
  return `${PARAGRAPH_CLASS_PREFIX}${id}`;
}

function runId(id: string): string {
  return `${RUN_ID_PREFIX}${id}`;
}

function fieldTextClass(id: string): string {
  return `${FIELD_TEXT_CLASS_PREFIX}${id}`;
}

function appendClass(node: ModelNode, className: string): void {
  node.className = `${node.className ?? ""} ${className}`.trim();
}

function descendants(
  node: ModelNode,
  predicate: (candidate: ModelNode) => boolean,
): Array<{ node: ModelNode; parent: ModelNode | null }> {
  const result: Array<{ node: ModelNode; parent: ModelNode | null }> = [];
  const visit = (candidate: ModelNode, parent: ModelNode | null) => {
    if (predicate(candidate)) result.push({ node: candidate, parent });
    candidate.children?.forEach((child) => visit(child, candidate));
  };
  visit(node, null);
  return result;
}

/**
 * Match editable manifest paragraphs to docx-preview's parsed run model once.
 * PAGE/NUMPAGES result runs stay renderer-owned, while surrounding Word runs
 * retain their stable manifest IDs and remain live-editable.
 */
export function markLiveDocxModel(
  documentModel: ModelDocument,
  manifest: DocxManifest,
  fieldTokens: ReadonlySet<string>,
): ReadonlySet<string> {
  const supported = new Set<string>();
  const modelParagraphs = descendants(
    documentModel.documentPart.body,
    (node) => node.type === "paragraph",
  );

  modelParagraphs.forEach(({ node }, index) => {
    const descriptor = manifest.paragraphs[index];
    if (!descriptor || descriptor.editable === false) return;
    const textIds = descriptor.segments.flatMap((segment) =>
      "textId" in segment ? [segment.textId] : [],
    );
    const textNodes = descendants(node, (candidate) => candidate.type === "text");
    if (textNodes.some(({ parent }) => !parent || parent.type !== "run")) return;
    const fieldIndexes = textNodes.flatMap(({ node: text }, textIndex) =>
      fieldTokens.has(text.text ?? "") ? [textIndex] : [],
    );
    const assignments = new Map<number, string>();
    const ignoredTextIds = new Set<string>();
    if (textNodes.length === textIds.length) {
      textNodes.forEach(({ node: text }, textIndex) => {
        const textId = textIds[textIndex];
        if (fieldTokens.has(text.text ?? "")) ignoredTextIds.add(textId);
        else assignments.set(textIndex, textId);
      });
    } else if (fieldIndexes.length === 1 && textNodes.length < textIds.length) {
      const fieldIndex = fieldIndexes[0];
      const suffixLength = textNodes.length - fieldIndex - 1;
      for (let textIndex = 0; textIndex < fieldIndex; textIndex += 1) {
        assignments.set(textIndex, textIds[textIndex]);
      }
      for (let offset = 0; offset < suffixLength; offset += 1) {
        assignments.set(
          fieldIndex + 1 + offset,
          textIds[textIds.length - suffixLength + offset],
        );
      }
      textIds
        .slice(fieldIndex, textIds.length - suffixLength)
        .forEach((textId) => ignoredTextIds.add(textId));
    } else {
      return;
    }

    appendClass(node, paragraphClass(descriptor.id));
    ignoredTextIds.forEach((textId) => appendClass(node, fieldTextClass(textId)));
    textNodes.forEach(({ parent }, textIndex) => {
      const textId = assignments.get(textIndex);
      if (textId) parent!.id = runId(textId);
    });
    supported.add(descriptor.id);
  });
  return supported;
}

function entityMap(snapshot: LiveDocxSnapshot): Map<string, ProjectedEntity> {
  return new Map(snapshot.entities.map((entity) => [entity.id, entity]));
}

function renderLiveText(
  run: HTMLElement,
  text: LiveDocxParagraph["texts"][number],
  entities: ReadonlyMap<string, ProjectedEntity>,
  insertionFormats: ReadonlyMap<string, InsertionFormat>,
  revisionMode: DocxRevisionMode,
): void {
  const owner = run.ownerDocument;
  const fragment = owner.createDocumentFragment();
  const boundaries = new Set<number>([0, text.value.length]);
  for (const annotation of text.annotations) {
    boundaries.add(annotation.start);
    boundaries.add(annotation.end);
  }
  const ordered = [...boundaries].sort((left, right) => left - right);

  const appendReferences = (offset: number) => {
    for (const annotation of text.annotations) {
      if (!annotation.last || annotation.end !== offset) continue;
      const entity = entities.get(annotation.entityId);
      if (!entity || entity.status !== "open") continue;
      if (annotation.kind === "redline") {
        const content = owner.createElement("span");
        const format = insertionFormats.get(annotation.entityId);
        if (format) {
          content.className = format.className;
          content.style.cssText = format.cssText;
        }
        content.textContent = entity.replacement ?? "";
        if (revisionMode === "final") {
          fragment.append(content);
        } else {
          const insertion = owner.createElement("ins");
          insertion.dataset.docxEntityId = annotation.entityId;
          insertion.append(content);
          fragment.append(insertion);
        }
      }
    }
  };

  appendReferences(0);
  for (let index = 0; index < ordered.length - 1; index += 1) {
    const start = ordered[index];
    const end = ordered[index + 1];
    const active = text.annotations.filter(
      (annotation) => annotation.start <= start && annotation.end >= end,
    );
    const redline = active.find(
      (annotation) => annotation.kind === "redline",
    );
    if (revisionMode === "final" && redline) {
      appendReferences(end);
      continue;
    }
    let content: Node = owner.createTextNode(text.value.slice(start, end));
    if (redline) {
      const deletion = owner.createElement("del");
      deletion.dataset.docxEntityId = redline.entityId;
      deletion.append(content);
      content = deletion;
    }
    const comments =
      revisionMode === "review"
        ? active.filter((annotation) => annotation.kind === "comment")
        : [];
    if (comments.length) {
      const highlight = owner.createElement("span");
      highlight.className = "docx-live-comment";
      highlight.dataset.docxCommentIds = JSON.stringify(
        comments.map((annotation) => annotation.entityId),
      );
      highlight.append(content);
      content = highlight;
    }
    fragment.append(content);
    appendReferences(end);
  }
  run.replaceChildren(fragment);
}

function patchParagraph(
  root: ParentNode,
  paragraph: LiveDocxParagraph,
  entities: ReadonlyMap<string, ProjectedEntity>,
  inheritedInsertionFormats: ReadonlyMap<string, InsertionFormat>,
  revisionMode: DocxRevisionMode,
): HTMLElement {
  const element = root.querySelector<HTMLElement>(
    `.${paragraphClass(paragraph.id)}`,
  );
  if (!element) {
    throw new Error(
      `Live update for ${paragraph.id} changes unsupported Word structure.`,
    );
  }
  const insertionFormats = new Map(inheritedInsertionFormats);
  for (const text of paragraph.texts) {
    const run = element.querySelector<HTMLElement>(`[id="${runId(text.id)}"]`);
    if (!run) continue;
    for (const annotation of text.annotations) {
      if (annotation.kind === "redline" && annotation.first) {
        insertionFormats.set(annotation.entityId, {
          className: run.className,
          cssText: run.style.cssText,
        });
      }
    }
  }
  for (const text of paragraph.texts) {
    const run = element.querySelector<HTMLElement>(`[id="${runId(text.id)}"]`);
    if (!run) {
      if (element.classList.contains(fieldTextClass(text.id))) continue;
      throw new Error(
        `Live update for ${paragraph.id} cannot map Word text ${text.id}.`,
      );
    }
    renderLiveText(run, text, entities, insertionFormats, revisionMode);
  }
  return element;
}

/** Cover the subscribe/getSnapshot startup gap, including entity-only payloads. */
export function liveCatchUpParagraphIds(
  applied: LiveDocxSnapshot,
  latest: LiveDocxSnapshot,
  supportedParagraphs: ReadonlySet<string>,
): ReadonlySet<string> {
  if (latest.revision <= applied.revision) return new Set();
  return new Set(
    [...latest.paragraphs.keys()].filter((id) => supportedParagraphs.has(id)),
  );
}

function collectInsertionFormats(
  roots: readonly ParentNode[],
  snapshot: LiveDocxSnapshot,
): Map<string, InsertionFormat> {
  const result = new Map<string, InsertionFormat>();
  for (const paragraph of snapshot.paragraphs.values()) {
    const firstAnnotations = paragraph.texts.flatMap((text) =>
      text.annotations
        .filter(
          (annotation) => annotation.kind === "redline" && annotation.first,
        )
        .map((annotation) => ({ annotation, text })),
    );
    if (!firstAnnotations.length) continue;
    const paragraphElement = roots
      .map((root) =>
        root.querySelector<HTMLElement>(`.${paragraphClass(paragraph.id)}`),
      )
      .find((element) => element !== null);
    if (!paragraphElement) continue;
    for (const { annotation, text } of firstAnnotations) {
      const run = paragraphElement.querySelector<HTMLElement>(
        `[id="${runId(text.id)}"]`,
      );
      if (run) {
        result.set(annotation.entityId, {
          className: run.className,
          cssText: run.style.cssText,
        });
      }
    }
  }
  return result;
}

export function applyInitialLiveSnapshot(
  container: HTMLElement,
  snapshot: LiveDocxSnapshot,
  supportedParagraphs: ReadonlySet<string>,
  revisionMode: DocxRevisionMode,
): void {
  const entities = entityMap(snapshot);
  const insertionFormats = collectInsertionFormats([container], snapshot);
  for (const [id, paragraph] of snapshot.paragraphs) {
    if (supportedParagraphs.has(id)) {
      patchParagraph(
        container,
        paragraph,
        entities,
        insertionFormats,
        revisionMode,
      );
    }
  }
  container.dataset.docxLiveRevision = String(snapshot.revision);
}

function directFlowBlocks(section: HTMLElement): HTMLElement[] {
  const header = section.querySelector(":scope > header");
  const footer = section.querySelector(":scope > footer");
  return [...section.children].flatMap((child) => {
    if (
      child === header ||
      child === footer ||
      (child as HTMLElement).hasAttribute("data-docx-page-decoration")
    ) {
      return [];
    }
    return child.tagName === "ARTICLE"
      ? ([...child.children] as HTMLElement[])
      : [child as HTMLElement];
  });
}

function flowIndex(element: Element): number | undefined {
  const value = Number.parseInt(
    element.getAttribute("data-docx-flow-index") ?? "",
    10,
  );
  return Number.isInteger(value) ? value : undefined;
}

function trimSectionBefore(section: HTMLElement, startFlow: number): void {
  for (const block of directFlowBlocks(section)) {
    const index = flowIndex(block);
    if (index !== undefined && index < startFlow) block.remove();
  }
  section.querySelectorAll(":scope > article").forEach((article) => {
    if (!article.children.length && !(article.textContent ?? "").trim()) {
      article.remove();
    }
  });
}

type SectionUpdate = {
  canonical: HTMLElement;
  index: number;
  oldPages: HTMLElement[];
  stageRoot: HTMLElement;
  stagedPages: HTMLElement[];
};

export type PreparedLiveDocxUpdate = {
  readonly revision: number;
  readonly stagedRoots: readonly HTMLElement[];
  commit(): number;
  discard(): void;
};

export class IncrementalDocxLayout {
  private canonicalSections: HTMLElement[];

  constructor(
    private readonly container: HTMLElement,
    private readonly revisionMode: DocxRevisionMode,
  ) {
    const sections = [
      ...container.querySelectorAll<HTMLElement>(
        ".docx-wrapper > section.docx",
      ),
    ];
    sections.forEach((section, sectionIndex) => {
      section.dataset.docxSectionIndex = String(sectionIndex);
      directFlowBlocks(section).forEach((block, index) => {
        block.dataset.docxFlowIndex = String(index);
      });
    });
    this.canonicalSections = sections.map(
      (section) => section.cloneNode(true) as HTMLElement,
    );
  }

  async prepare(
    snapshot: LiveDocxSnapshot,
    changedParagraphIds: ReadonlySet<string>,
    stageParent: ShadowRoot,
    isCurrent: () => boolean,
    layout: (container: HTMLElement) => void,
  ): Promise<PreparedLiveDocxUpdate> {
    const entities = entityMap(snapshot);
    const insertionFormats = collectInsertionFormats(
      this.canonicalSections,
      snapshot,
    );
    const candidates = new Map<number, HTMLElement>();
    const changedFlows = new Map<number, Set<number>>();

    for (const paragraphId of changedParagraphIds) {
      const paragraph = snapshot.paragraphs.get(paragraphId);
      if (!paragraph) {
        throw new Error(`Live update is missing paragraph ${paragraphId}.`);
      }
      let sectionIndex = this.canonicalSections.findIndex((section) =>
        section.querySelector(`.${paragraphClass(paragraphId)}`),
      );
      if (sectionIndex < 0) {
        throw new Error(
          `Live update for ${paragraphId} changes unsupported Word structure.`,
        );
      }
      let candidate = candidates.get(sectionIndex);
      if (!candidate) {
        candidate = this.canonicalSections[sectionIndex].cloneNode(
          true,
        ) as HTMLElement;
        candidates.set(sectionIndex, candidate);
      }
      const changed = patchParagraph(
        candidate,
        paragraph,
        entities,
        insertionFormats,
        this.revisionMode,
      );
      const block = changed.closest<HTMLElement>("[data-docx-flow-index]");
      const index = block ? flowIndex(block) : undefined;
      if (index === undefined) {
        throw new Error(`Live update cannot locate ${paragraphId} in page flow.`);
      }
      const flows = changedFlows.get(sectionIndex) ?? new Set<number>();
      flows.add(index);
      changedFlows.set(sectionIndex, flows);
    }

    const updates: SectionUpdate[] = [];
    const stageRoots: HTMLElement[] = [];
    try {
      for (const [sectionIndex, canonical] of candidates) {
        if (!isCurrent()) throw new Error("Live update was superseded.");
        const oldPages = [
          ...this.container.querySelectorAll<HTMLElement>(
            `.docx-wrapper > section.docx[data-docx-section-index="${sectionIndex}"]`,
          ),
        ];
        if (!oldPages.length) {
          throw new Error(`Live update cannot locate Word section ${sectionIndex}.`);
        }
        const changed = changedFlows.get(sectionIndex)!;
        let pageIndex = oldPages.findIndex((page) =>
          [...page.querySelectorAll("[data-docx-flow-index]")].some((block) => {
            const index = flowIndex(block);
            return index !== undefined && changed.has(index);
          }),
        );
        if (pageIndex < 0) pageIndex = 0;
        let startFlow = Math.min(...changed);
        const firstChangedPageFlows = [
          ...oldPages[pageIndex].querySelectorAll("[data-docx-flow-index]"),
        ]
          .map(flowIndex)
          .filter((value): value is number => value !== undefined);
        if (
          pageIndex > 0 &&
          firstChangedPageFlows.some(
            (index) => changed.has(index) && index === Math.min(...firstChangedPageFlows),
          )
        ) {
          // A shorter first block can pull content back across the existing
          // page boundary. Include the preceding page and its keep-next chain.
          pageIndex -= 1;
        }
        while (true) {
          const pageFlows = [
            ...oldPages[pageIndex].querySelectorAll("[data-docx-flow-index]"),
          ]
            .map(flowIndex)
            .filter((value): value is number => value !== undefined);
          startFlow = Math.min(startFlow, ...pageFlows);
          const firstPage = oldPages.findIndex((page) =>
            page.querySelector(
              `[data-docx-flow-index="${startFlow}"]`,
            ),
          );
          if (firstPage < 0 || firstPage === pageIndex) break;
          pageIndex = firstPage;
        }

        const stagedSection = canonical.cloneNode(true) as HTMLElement;
        trimSectionBefore(stagedSection, startFlow);
        const stageRoot = document.createElement("div");
        stageRoot.className = "docx-render-target docx-incremental-stage";
        stageRoot.dataset.docxRevisionMode = this.revisionMode;
        const wrapper = document.createElement("div");
        wrapper.className = "docx-wrapper";
        wrapper.append(stagedSection);
        stageRoot.append(wrapper);
        stageParent.append(stageRoot);
        stageRoots.push(stageRoot);
        await Promise.all([
          document.fonts.ready,
          ...[...stageRoot.querySelectorAll("img")].map((image) =>
            image.decode().catch(() => undefined),
          ),
        ]);
        layout(stageRoot);
        await paginateDocx(stageRoot, isCurrent, { restoreHighlights: false });
        if (!isCurrent()) throw new Error("Live update was superseded.");
        layout(stageRoot);
        updates.push({
          canonical,
          index: sectionIndex,
          oldPages: oldPages.slice(pageIndex),
          stageRoot,
          stagedPages: [
            ...wrapper.querySelectorAll<HTMLElement>(":scope > section.docx"),
          ],
        });
      }
    } catch (error) {
      stageRoots.forEach((stageRoot) => stageRoot.remove());
      throw error;
    }

    let finished = false;
    return {
      revision: snapshot.revision,
      stagedRoots: updates.map((update) => update.stageRoot),
      commit: () => {
        if (finished) {
          return this.container.querySelectorAll(
            ".docx-wrapper > section.docx",
          ).length;
        }
        finished = true;
        for (const update of updates) {
          const fragment = document.createDocumentFragment();
          update.stagedPages.forEach((page) => fragment.append(page));
          update.oldPages[0].before(fragment);
          update.oldPages.forEach((page) => page.remove());
          update.stageRoot.remove();
          this.canonicalSections[update.index] = update.canonical;
        }
        this.container.dataset.docxLiveRevision = String(snapshot.revision);
        restoreCommentHighlights(this.container);
        return this.container.querySelectorAll(
          ".docx-wrapper > section.docx",
        ).length;
      },
      discard: () => {
        if (finished) return;
        finished = true;
        updates.forEach((update) => update.stageRoot.remove());
      },
    };
  }
}
