import {
  useEffect,
  useMemo,
  useRef,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from "react";
import type { ProjectedEntity, ReviewItem, ReviewItemRenderContext } from "./types";
import type { ReviewSelectionRequest } from "./renderers/reviewNavigation";

type ReviewPanelProps = {
  items: readonly ReviewItem[];
  entities: readonly ProjectedEntity[];
  projectionReady: boolean;
  selection: ReviewSelectionRequest | null;
  navigationEnabled: boolean;
  renderItem?: (context: ReviewItemRenderContext) => ReactNode;
  onSelectEntity: (entityId: string) => void;
};

const INTERACTIVE_REVIEW_CONTENT =
  "button, a, input, select, textarea, summary, [contenteditable], [data-review-action]";

function hasInteractiveTarget(target: EventTarget, currentTarget: HTMLElement) {
  const interactive =
    target instanceof Element ? target.closest(INTERACTIVE_REVIEW_CONTENT) : null;
  return interactive !== null && interactive !== currentTarget;
}

function diagnosticFor(entity: ProjectedEntity | undefined, projectionReady: boolean) {
  if (!entity) return projectionReady ? "The renderer did not return this review item." : null;
  if (entity.status === "invalid" || entity.status === "conflict") {
    return entity.reason ?? `This ${entity.kind} has a ${entity.status} anchor.`;
  }
  return entity.reason ?? null;
}

export function ReviewPanel({
  items,
  entities,
  projectionReady,
  selection,
  navigationEnabled,
  renderItem,
  onSelectEntity,
}: ReviewPanelProps) {
  const panelRef = useRef<HTMLElement>(null);
  const itemRefs = useRef(new Map<string, HTMLElement>());
  const entityById = useMemo(
    () => new Map(entities.map((entity) => [entity.id, entity])),
    [entities],
  );

  useEffect(() => {
    if (!selection || selection.origin !== "document" || !navigationEnabled) return;
    const panel = panelRef.current;
    const item = itemRefs.current.get(selection.entityId);
    if (!panel || !item) return;
    const panelBounds = panel.getBoundingClientRect();
    const itemBounds = item.getBoundingClientRect();
    if (itemBounds.top >= panelBounds.top + 12 && itemBounds.bottom <= panelBounds.bottom - 12) return;
    item.scrollIntoView({
      block: "nearest",
      behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches
        ? "auto"
        : "smooth",
    });
  }, [navigationEnabled, selection]);

  function keyDown(event: KeyboardEvent<HTMLElement>, entityId: string) {
    if (event.key !== "Enter" && event.key !== " ") return;
    if (hasInteractiveTarget(event.target, event.currentTarget)) return;
    event.preventDefault();
    onSelectEntity(entityId);
  }

  function click(event: MouseEvent<HTMLElement>, entityId: string) {
    if (hasInteractiveTarget(event.target, event.currentTarget)) return;
    const textSelection = window.getSelection();
    if (textSelection && !textSelection.isCollapsed) return;
    onSelectEntity(entityId);
  }

  return (
    <aside
      ref={panelRef}
      className="docx-preview-review-panel"
      aria-label="Review comments and redlines"
    >
      <div className="docx-preview-review-heading">
        <h2>Review</h2>
        <span>{items.length}</span>
      </div>
      <p className="docx-preview-review-description">
        Comments and proposed replacements in this document.
      </p>
      {items.length === 0 && <div className="docx-preview-review-empty">No review items.</div>}
      {items.map((item) => {
        const entity = entityById.get(item.id);
        const canNavigate = navigationEnabled && entity?.status === "open";
        const selected = canNavigate && selection?.entityId === item.id;
        const diagnostic = diagnosticFor(entity, projectionReady);
        return (
          <article
            className={`docx-preview-review-item docx-preview-review-item--${item.kind}${selected ? " is-selected" : ""}`}
            key={item.id}
            data-review-item-id={item.id}
            ref={(element) => {
              if (element) itemRefs.current.set(item.id, element);
              else itemRefs.current.delete(item.id);
            }}
            role={canNavigate ? "button" : undefined}
            tabIndex={canNavigate ? 0 : undefined}
            aria-pressed={canNavigate ? selected : undefined}
            aria-label={canNavigate ? `Show ${item.kind} ${item.id} in the document` : undefined}
            onClick={canNavigate ? (event) => click(event, item.id) : undefined}
            onKeyDown={canNavigate ? (event) => keyDown(event, item.id) : undefined}
          >
            <div className="docx-preview-review-item-heading">
              <span>{item.kind === "redline" ? "Redline" : "Comment"}</span>
              <span>{entity?.status ?? "pending"}</span>
            </div>
            {renderItem ? (
              renderItem({ item, entity: entity ?? null, selected, diagnostic })
            ) : (
              <>
                <blockquote>{item.anchor.quote}</blockquote>
                {item.kind === "redline" && (
                  <p className="docx-preview-replacement-text">
                    {item.replacement || "Delete selected text"}
                  </p>
                )}
                {item.body && <p>{item.body}</p>}
              </>
            )}
            {diagnostic && (
              <p className="docx-preview-review-warning" data-review-diagnostic>
                {diagnostic}
              </p>
            )}
          </article>
        );
      })}
    </aside>
  );
}
