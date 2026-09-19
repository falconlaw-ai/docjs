export type ReviewSelectionRequest = {
  entityId: string;
  origin: "document" | "review";
  sequence: number;
};

const ENTITY_SELECTOR =
  "[data-docx-entity-id], [data-docx-comment-ids]";
const SELECTED_CLASS = "docx-review-selected";

function commentEntityIds(element: HTMLElement): string[] {
  const encoded = element.dataset.docxCommentIds;
  if (!encoded) return [];
  try {
    const value: unknown = JSON.parse(encoded);
    return Array.isArray(value)
      ? value.filter((id): id is string => typeof id === "string")
      : [];
  } catch {
    // Older live DOM used a space-separated value. Treat it as one ID rather
    // than splitting a valid entity ID that happens to contain spaces.
    return [encoded];
  }
}

function entityIds(element: HTMLElement): string[] {
  const exact = element.dataset.docxEntityId;
  return [...(exact === undefined ? [] : [exact]), ...commentEntityIds(element)];
}

function entityTargets(container: HTMLElement, entityId: string): HTMLElement[] {
  return [...container.querySelectorAll<HTMLElement>(ENTITY_SELECTOR)].filter(
    (element) => entityIds(element).includes(entityId),
  );
}

function activationEntityIds(
  container: HTMLElement,
  target: EventTarget | null,
): string[] {
  if (!(target instanceof Node)) return [];
  const ids: string[] = [];
  let element =
    target instanceof HTMLElement ? target : target.parentElement;
  while (element && element !== container) {
    if (element.matches(ENTITY_SELECTOR)) {
      for (const id of entityIds(element)) {
        if (!ids.includes(id)) ids.push(id);
      }
    }
    element = element.parentElement;
  }
  return ids;
}

function selectionTouches(
  selection: Selection | null,
  container: HTMLElement,
): boolean {
  if (!selection || selection.isCollapsed) return false;
  return [selection.anchorNode, selection.focusNode].some(
    (node) => node !== null && container.contains(node),
  );
}

function hasDocumentTextSelection(container: HTMLElement): boolean {
  const root = container.getRootNode();
  const shadowSelection =
    root instanceof ShadowRoot
      ? (
          root as ShadowRoot & {
            getSelection?: () => Selection | null;
          }
        ).getSelection?.()
      : null;
  return (
    selectionTouches(shadowSelection ?? null, container) ||
    selectionTouches(container.ownerDocument.getSelection(), container)
  );
}

function preferredScrollTarget(targets: readonly HTMLElement[]): HTMLElement {
  return (
    targets.find((target) => !target.closest("[data-docx-repeated]")) ??
    targets[0]
  );
}

function scrollWithinViewport(
  viewport: HTMLElement,
  target: HTMLElement,
): void {
  const viewportRect = viewport.getBoundingClientRect();
  const targetRect = target.getBoundingClientRect();
  const top = Math.max(
    0,
    Math.min(
      viewport.scrollHeight - viewport.clientHeight,
      viewport.scrollTop +
        targetRect.top -
        viewportRect.top -
        (viewport.clientHeight - targetRect.height) / 2,
    ),
  );
  const left = Math.max(
    0,
    Math.min(
      viewport.scrollWidth - viewport.clientWidth,
      viewport.scrollLeft +
        targetRect.left -
        viewportRect.left -
        (viewport.clientWidth - targetRect.width) / 2,
    ),
  );
  const behavior = viewport.ownerDocument.defaultView?.matchMedia(
    "(prefers-reduced-motion: reduce)",
  ).matches
    ? "auto"
    : "smooth";
  viewport.scrollTo({ behavior, left, top });
}

function requestKey(selection: ReviewSelectionRequest): string {
  return `${selection.sequence}:${selection.entityId}`;
}

export type ReviewNavigationController = {
  dispose(): void;
  refresh(): void;
  setSelection(selection: ReviewSelectionRequest | null): void;
};

export function createReviewNavigation(
  container: HTMLElement,
  scrollViewport: HTMLElement | null,
  onSelectEntity: (entityId: string) => void,
): ReviewNavigationController {
  let current: ReviewSelectionRequest | null = null;
  let disposed = false;
  let lastScrolledKey: string | null = null;
  let pendingScrollKey: string | null = null;
  let scrollFrame = 0;

  const enhanceTargets = () => {
    const keyboardCommentIds = new Set<string>();
    for (const target of container.querySelectorAll<HTMLElement>(ENTITY_SELECTOR)) {
      target.classList.add("docx-review-target");
      const commentIds = commentEntityIds(target);
      const isCommentText = commentIds.length > 0;
      const isKeyboardTarget =
        !isCommentText || commentIds.some((id) => !keyboardCommentIds.has(id));
      commentIds.forEach((id) => keyboardCommentIds.add(id));
      if (isKeyboardTarget) {
        target.tabIndex = 0;
        target.setAttribute("role", "button");
        if (!target.hasAttribute("aria-pressed")) {
          target.setAttribute("aria-pressed", "false");
        }
        if (!target.hasAttribute("aria-label")) {
          target.setAttribute(
            "aria-label",
            isCommentText ? "Review comment" : "Review redline",
          );
        }
      } else {
        target.removeAttribute("tabindex");
        target.removeAttribute("role");
        target.removeAttribute("aria-pressed");
        target.removeAttribute("aria-label");
      }
    }
  };

  const applySelectedClass = (): HTMLElement[] => {
    container
      .querySelectorAll<HTMLElement>(`.${SELECTED_CLASS}`)
      .forEach((element) => {
        element.classList.remove(SELECTED_CLASS);
        if (element.getAttribute("role") === "button") {
          element.setAttribute("aria-pressed", "false");
        }
      });
    if (!current) return [];
    const targets = entityTargets(container, current.entityId);
    targets.forEach((target) => {
      target.classList.add(SELECTED_CLASS);
      if (target.getAttribute("role") === "button") {
        target.setAttribute("aria-pressed", "true");
      }
    });
    return targets;
  };

  const schedulePendingScroll = (targets: readonly HTMLElement[]) => {
    if (
      disposed ||
      !scrollViewport ||
      !current ||
      !pendingScrollKey ||
      !targets.length
    ) {
      return;
    }
    const scheduledKey = pendingScrollKey;
    cancelAnimationFrame(scrollFrame);
    scrollFrame = requestAnimationFrame(() => {
      // Incremental layout restores the reading position on its next frame.
      // Run after that restore so a pending review request remains authoritative.
      scrollFrame = requestAnimationFrame(() => {
        if (
          disposed ||
          !current ||
          current.origin !== "review" ||
          pendingScrollKey !== scheduledKey ||
          requestKey(current) !== scheduledKey
        ) {
          return;
        }
        const currentTargets = entityTargets(container, current.entityId);
        if (!currentTargets.length) return;
        scrollWithinViewport(
          scrollViewport,
          preferredScrollTarget(currentTargets),
        );
        lastScrolledKey = scheduledKey;
        pendingScrollKey = null;
      });
    });
  };

  const activate = (ids: readonly string[]) => {
    if (!ids.length) return;
    const selectedIndex = current ? ids.indexOf(current.entityId) : -1;
    const entityId = ids[(selectedIndex + 1) % ids.length];
    current = {
      entityId,
      origin: "document",
      sequence: current?.sequence ?? 0,
    };
    pendingScrollKey = null;
    applySelectedClass();
    onSelectEntity(entityId);
  };

  const onClick = (event: MouseEvent) => {
    const ids = activationEntityIds(container, event.target);
    if (!ids.length || hasDocumentTextSelection(container)) return;
    activate(ids);
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const ids = activationEntityIds(container, event.target);
    if (!ids.length) return;
    event.preventDefault();
    activate(ids);
  };

  container.addEventListener("click", onClick);
  container.addEventListener("keydown", onKeyDown);
  enhanceTargets();

  return {
    dispose: () => {
      disposed = true;
      cancelAnimationFrame(scrollFrame);
      container.removeEventListener("click", onClick);
      container.removeEventListener("keydown", onKeyDown);
    },
    refresh: () => {
      enhanceTargets();
      schedulePendingScroll(applySelectedClass());
    },
    setSelection: (selection) => {
      current = selection;
      const key = selection?.origin === "review" ? requestKey(selection) : null;
      pendingScrollKey = key !== lastScrolledKey ? key : null;
      schedulePendingScroll(applySelectedClass());
    },
  };
}
