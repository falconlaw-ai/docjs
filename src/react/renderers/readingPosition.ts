export type ReadingPosition = {
  progress: number;
  text: string;
  offset: number;
};

function blocks(viewport: HTMLElement) {
  const custom = viewport.querySelector(".docx-preview")?.shadowRoot;
  return Array.from(
    (custom ?? viewport).querySelectorAll<HTMLElement>(
      custom ? "section.docx > article p" : ".superdoc-page .superdoc-fragment",
    ),
  ).filter(
    (element) =>
      !element.closest(".superdoc-page-header, .superdoc-page-footer"),
  );
}

const normalized = (element: HTMLElement) =>
  // The painters disagree about spaces around list labels and between lines.
  (element.textContent ?? "")
    .replace(/[\s\u200b-\u200d\ufeff\u00ad]+/g, "")
    .replace(/^(?:\d+(?:\.\d+)*[.)]?|\([a-z]\)|[•▪])/, "");

export function captureReadingPosition(viewport: HTMLElement): ReadingPosition {
  const top = viewport.getBoundingClientRect().top;
  const block = blocks(viewport)
    .map((element) => ({ element, bounds: element.getBoundingClientRect() }))
    .filter(
      ({ element, bounds }) =>
        bounds.bottom > top + 10 &&
        bounds.top < top + viewport.clientHeight &&
        normalized(element).length >= 24,
    )
    .sort((a, b) => a.bounds.top - b.bounds.top)[0]?.element;
  return {
    progress:
      viewport.scrollTop /
      Math.max(1, viewport.scrollHeight - viewport.clientHeight),
    text: block ? normalized(block).slice(0, 100) : "",
    offset: block ? block.getBoundingClientRect().top - top : 0,
  };
}

export function restoreReadingPosition(
  viewport: HTMLElement,
  position: ReadingPosition,
  matchText = false,
) {
  if (matchText && position.text) {
    const candidates = blocks(viewport).filter((element) =>
      normalized(element).includes(position.text),
    );
    // Repeated headings are ambiguous; keep the proportional position in that case.
    if (candidates.length === 1) {
      viewport.scrollTop +=
        candidates[0].getBoundingClientRect().top -
        viewport.getBoundingClientRect().top -
        position.offset;
      return;
    }
  }
  if (!matchText)
    viewport.scrollTop =
      position.progress *
      Math.max(0, viewport.scrollHeight - viewport.clientHeight);
}
