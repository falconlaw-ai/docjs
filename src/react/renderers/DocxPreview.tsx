import { useEffect, useRef, useState } from "react";
import { parseAsync, renderDocument, type Options } from "../../docx-preview";
import {
  createItemSource,
  type ItemDocxSource,
} from "../../review/itemSource";
import type {
  LiveDocxChange,
  ProjectedEntity,
  ReviewItem,
} from "../../review/model";
import { applyPaginationHints, paginateDocx } from "./paginateDocx";
import {
  markPageFields,
  preparePageFields,
  updatePageFields,
} from "./docxFields";
import { applyDocxGeometry, prepareDocxGeometry } from "./docxGeometry";
import {
  normalizeDocxTypography,
  createDocxNumberingFixStyles,
  createDocxHyphenationStyles,
  layoutDocxTabs,
} from "./docxTypography";
import {
  applyInitialLiveSnapshot,
  IncrementalDocxLayout,
  liveCatchUpParagraphIds,
  markLiveDocxModel,
  type DocxRevisionMode,
  type PreparedLiveDocxUpdate,
} from "./incrementalDocx";
import {
  captureReadingPosition,
  restoreReadingPosition,
} from "./readingPosition";
import {
  createReviewNavigation,
  type ReviewNavigationController,
  type ReviewSelectionRequest,
} from "./reviewNavigation";

export type { ReviewSelectionRequest } from "./reviewNavigation";
export type { DocxRevisionMode } from "./incrementalDocx";

import previewStyles from "./docx-preview.css";

export type DocxPreviewProps = {
  file: File;
  items?: readonly ReviewItem[];
  revisionMode?: DocxRevisionMode;
  zoom?: number;
  onPageWidth?: (width: number) => void;
  onLoad?: (pages: number) => void;
  onError?: (message: string) => void;
  onRefreshError?: (message: string) => void;
  selection?: ReviewSelectionRequest | null;
  onSelectEntity?: (entityId: string) => void;
  onEntitiesChange?: (entities: readonly ProjectedEntity[]) => void;
};

type ItemSourceState =
  | { file: File; source: ItemDocxSource }
  | { error: string; file: File };

const RENDER_OPTIONS = {
  breakPages: true,
  className: "docx",
  debug: false,
  experimental: false,
  ignoreFonts: false,
  ignoreHeight: false,
  // Geometry neutralizes Word's cached layout markers. Keeping this false lets
  // docx-preview group continuous sections instead of forcing each onto a page.
  ignoreLastRenderedPageBreak: false,
  ignoreWidth: false,
  inWrapper: true,
  hideWrapperOnPrint: false,
  renderAltChunks: false,
  renderChanges: true,
  renderComments: true,
  renderEndnotes: true,
  renderFooters: true,
  renderFootnotes: true,
  renderHeaders: true,
  trimXmlDeclaration: true,
  // Keeping resources as data URLs avoids object URL lifetime leaks when files change.
  useBase64URL: true,
} satisfies Partial<Options>;

function renderOptions(revisionMode: DocxRevisionMode): Partial<Options> {
  return revisionMode === "final"
    ? { ...RENDER_OPTIONS, renderChanges: false, renderComments: false }
    : RENDER_OPTIONS;
}

const UNSAFE_ELEMENTS = "script, iframe, object, embed, link, meta, base";
const RESOURCE_ATTRIBUTES = ["src", "poster"] as const;
const LEGACY_LIST_MARKERS = new Map([
  ["\uf0b7", "•"],
  ["\uf0a7", "▪"],
]);

function toErrorMessage(error: unknown): string {
  return error instanceof Error && error.message
    ? error.message
    : "This Word document couldn't be previewed.";
}

function isLocalResource(value: string): boolean {
  const normalized = value.trim().toLowerCase();
  return (
    normalized.startsWith("data:") ||
    normalized.startsWith("blob:") ||
    normalized.startsWith("#")
  );
}

function isolateRenderedContent(container: HTMLElement): void {
  container
    .querySelectorAll(UNSAFE_ELEMENTS)
    .forEach((element) => element.remove());

  for (const attribute of RESOURCE_ATTRIBUTES) {
    container
      .querySelectorAll<HTMLElement>(`[${attribute}]`)
      .forEach((element) => {
        const value = element.getAttribute(attribute);
        if (value && !isLocalResource(value)) {
          element.removeAttribute(attribute);
        }
      });
  }

  container.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((anchor) => {
    const href = anchor.getAttribute("href")?.trim() ?? "";
    if (/^javascript:/i.test(href) || /^data:/i.test(href)) {
      anchor.removeAttribute("href");
      return;
    }

    if (/^https?:/i.test(href)) {
      anchor.target = "_blank";
      anchor.rel = "noopener noreferrer";
    }
  });
}

function normalizeLegacyListMarkers(container: HTMLElement): void {
  container
    .querySelectorAll<HTMLParagraphElement>('p[class*="docx-num-"]')
    .forEach((paragraph) => {
      const content = getComputedStyle(paragraph, "::before").content;
      for (const [legacyMarker, marker] of LEGACY_LIST_MARKERS) {
        if (content.includes(legacyMarker)) {
          paragraph.dataset.docxListMarker = marker;
          break;
        }
      }
    });
}

function createPreviewFrame(measurementWidth: number) {
  const measurementHost = document.createElement("div");
  measurementHost.className = "docx-measurement-host";
  measurementHost.setAttribute("aria-hidden", "true");
  measurementHost.style.width = `${Math.max(1, Math.ceil(measurementWidth))}px`;

  const styleContainer = document.createElement("div");
  styleContainer.className = "docx-style-container";

  const viewport = document.createElement("div");
  viewport.className = "docx-viewport";

  const sizer = document.createElement("div");
  sizer.className = "docx-scale-sizer";

  const scaleLayer = document.createElement("div");
  scaleLayer.className = "docx-scale-layer";

  const renderTarget = document.createElement("div");
  renderTarget.className = "docx-render-target";

  scaleLayer.append(renderTarget);
  sizer.append(scaleLayer);
  viewport.append(sizer);

  const componentStyles = document.createElement("style");
  componentStyles.textContent = previewStyles;
  measurementHost
    .attachShadow({ mode: "open" })
    .append(styleContainer, componentStyles, viewport);

  return {
    measurementHost,
    renderTarget,
    scaleLayer,
    sizer,
    styleContainer,
    viewport,
  };
}

function getPreviewStyles(shadowRoot: ShadowRoot): HTMLStyleElement {
  const existing = shadowRoot.querySelector<HTMLStyleElement>(
    "style[data-docx-preview-styles]",
  );
  if (existing) {
    return existing;
  }

  const styles = document.createElement("style");
  styles.dataset.docxPreviewStyles = "";
  styles.textContent = previewStyles;
  shadowRoot.append(styles);
  return styles;
}

function fitRenderedPages(
  viewport: HTMLElement,
  sizer: HTMLElement,
  scaleLayer: HTMLElement,
  zoom?: number,
  onPageWidth?: (width: number) => void,
): void {
  // scrollWidth/scrollHeight are unscaled. Resetting the observed layout here
  // would create a ResizeObserver feedback loop on narrow viewports.
  const naturalWidth = Math.ceil(scaleLayer.scrollWidth);
  const naturalHeight = Math.ceil(scaleLayer.scrollHeight);
  const viewportStyle = getComputedStyle(viewport);
  const horizontalPadding =
    Number.parseFloat(viewportStyle.paddingLeft) +
    Number.parseFloat(viewportStyle.paddingRight);
  const availableWidth = Math.max(0, viewport.clientWidth - horizontalPadding);

  if (naturalWidth === 0 || naturalHeight === 0 || availableWidth === 0) {
    return;
  }

  onPageWidth?.(naturalWidth);
  const scale =
    zoom === undefined
      ? Math.max(0.25, Math.min(1, availableWidth / naturalWidth))
      : zoom / 100;
  scaleLayer.style.transform = `scale(${scale})`;
  sizer.style.width = `${Math.ceil(naturalWidth * scale)}px`;
  sizer.style.height = `${Math.ceil(naturalHeight * scale)}px`;
}

/**
 * A layout-oriented HTML preview. It uses stored page geometry and page-break
 * hints, but it does not run Word's layout engine, so pagination and complex
 * fields can differ from Word.
 */
export function DocxPreview({
  file,
  items,
  onLoad,
  onError,
  onRefreshError,
  zoom,
  onPageWidth,
  revisionMode = "review",
  selection,
  onSelectEntity,
  onEntitiesChange,
}: DocxPreviewProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [itemSourceState, setItemSourceState] =
    useState<ItemSourceState | null>(null);
  const renderSequenceRef = useRef(0);
  const onLoadRef = useRef(onLoad);
  const onErrorRef = useRef(onError);
  const onRefreshErrorRef = useRef(onRefreshError);
  const zoomRef = useRef(zoom);
  const widthCallback = useRef(onPageWidth);
  const selectionRef = useRef(selection ?? null);
  const onSelectEntityRef = useRef(onSelectEntity);
  const onEntitiesChangeRef = useRef(onEntitiesChange);
  const itemsRef = useRef(items);
  const fitRef = useRef<(() => void) | null>(null);
  const committedFrameRef = useRef<{
    dispose: () => void;
    file: File;
    fit: () => void;
    navigation: ReviewNavigationController;
    pageCount: number;
  } | null>(null);
  zoomRef.current = zoom;
  widthCallback.current = onPageWidth;
  selectionRef.current = selection ?? null;
  onSelectEntityRef.current = onSelectEntity;
  onEntitiesChangeRef.current = onEntitiesChange;
  itemsRef.current = items;
  const usesItems = items !== undefined;
  const source =
    usesItems &&
    itemSourceState?.file === file &&
    "source" in itemSourceState
      ? itemSourceState.source
      : undefined;
  const sourceError =
    usesItems &&
    itemSourceState?.file === file &&
    "error" in itemSourceState
      ? itemSourceState.error
      : undefined;
  const sourcePending = usesItems && source === undefined && !sourceError;

  useEffect(() => {
    let cancelled = false;
    let createdSource: ItemDocxSource | undefined;

    setItemSourceState(null);
    if (!usesItems) {
      return;
    }

    void createItemSource(file, itemsRef.current ?? []).then(
      (nextSource) => {
        if (cancelled) {
          nextSource.dispose();
          return;
        }
        createdSource = nextSource;
        nextSource.setItems(itemsRef.current ?? []);
        setItemSourceState({ file, source: nextSource });
      },
      (error) => {
        if (cancelled) return;
        const message = toErrorMessage(error);
        setItemSourceState({ error: message, file });
        const committed = committedFrameRef.current;
        if (committed?.file === file) {
          onRefreshErrorRef.current?.(message);
        } else {
          onErrorRef.current?.(message);
        }
      },
    );

    return () => {
      cancelled = true;
      createdSource?.dispose();
    };
  }, [file, usesItems]);

  useEffect(() => {
    if (source && items) {
      source.setItems(items);
    }
  }, [items, source]);

  useEffect(() => {
    fitRef.current?.();
  }, [zoom]);

  useEffect(() => {
    committedFrameRef.current?.navigation.setSelection(selection ?? null);
  }, [selection]);

  useEffect(
    () => () => {
      renderSequenceRef.current += 1;
      committedFrameRef.current?.dispose();
      committedFrameRef.current = null;
      fitRef.current = null;
    },
    [],
  );

  onLoadRef.current = onLoad;
  onErrorRef.current = onError;
  onRefreshErrorRef.current = onRefreshError;

  useEffect(() => {
    const host = hostRef.current;
    if (!host) {
      return;
    }

    const shadowRoot = host.shadowRoot ?? host.attachShadow({ mode: "open" });
    const previewStyleElement = getPreviewStyles(shadowRoot);
    const previousDocument = committedFrameRef.current;
    if (previousDocument && previousDocument.file !== file) {
      previousDocument.dispose();
      committedFrameRef.current = null;
      fitRef.current = null;
      shadowRoot.replaceChildren(previewStyleElement);
    }
    if (sourcePending || sourceError) {
      host.setAttribute("aria-busy", sourcePending ? "true" : "false");
      return;
    }
    const renderSequence = ++renderSequenceRef.current;
    let cancelled = false;
    let measurementHost: HTMLElement | null = null;
    let stopLiveUpdates = () => {};

    host.setAttribute("aria-busy", "true");

    const render = async () => {
      const frame = createPreviewFrame(host.getBoundingClientRect().width);
      const options = renderOptions(revisionMode);

      try {
        const buffer = await (source?.file ?? file).arrayBuffer();
        if (cancelled || renderSequence !== renderSequenceRef.current) {
          return;
        }

        // Render away from the live tree so a stale file cannot leave a half-built
        // document behind if the selected file changes while parsing.
        const fields = await preparePageFields(buffer);
        const documentModel = await parseAsync(fields.buffer, options);
        const liveParagraphs = source
          ? markLiveDocxModel(
              documentModel,
              source.manifest,
              new Set(Object.values(fields.tokens)),
            )
          : undefined;
        markPageFields(documentModel, fields.tokens);
        const geometry = await prepareDocxGeometry(
          fields.buffer,
          documentModel,
        );
        const typography = normalizeDocxTypography(documentModel);
        applyPaginationHints(documentModel);
        const nodes = await renderDocument(documentModel, options);
        for (const node of nodes) {
          (node.nodeName === "STYLE"
            ? frame.styleContainer
            : frame.renderTarget
          ).appendChild(node);
        }
        const documentStyles = document.createElement("style");
        documentStyles.textContent =
          createDocxNumberingFixStyles(documentModel) +
          "\n" +
          createDocxHyphenationStyles(documentModel);
        frame.styleContainer.append(documentStyles);

        if (cancelled || renderSequence !== renderSequenceRef.current) {
          return;
        }

        isolateRenderedContent(frame.renderTarget);
        frame.renderTarget.dataset.docxRevisionMode = revisionMode;

        measurementHost = frame.measurementHost;
        shadowRoot.append(measurementHost);
        applyDocxGeometry(frame.renderTarget, geometry);
        normalizeLegacyListMarkers(frame.renderTarget);
        const initialLiveSnapshot = source?.getSnapshot();
        if (initialLiveSnapshot && liveParagraphs) {
          applyInitialLiveSnapshot(
            frame.renderTarget,
            initialLiveSnapshot,
            liveParagraphs,
            revisionMode,
          );
        }

        // Pagination needs final font/image metrics at the document's natural
        // width. Viewport scaling happens only after page layout is complete.
        await Promise.all([
          document.fonts.ready,
          ...[...frame.renderTarget.querySelectorAll("img")].map((img) =>
            img.decode().catch(() => undefined),
          ),
        ]);
        const isCurrent = () =>
          !cancelled && renderSequence === renderSequenceRef.current;
        if (!isCurrent()) return;
        layoutDocxTabs(frame.renderTarget, typography);
        const incrementalLayout = source
          ? new IncrementalDocxLayout(frame.renderTarget, revisionMode)
          : undefined;
        const pageCount = await paginateDocx(frame.renderTarget, isCurrent);
        if (!isCurrent()) return;
        updatePageFields(frame.renderTarget);
        layoutDocxTabs(frame.renderTarget, typography);

        // Match the current host width and zoom before publishing so the new
        // document never appears at its natural size for a frame.
        frame.measurementHost.style.width = `${Math.max(
          1,
          Math.ceil(host.getBoundingClientRect().width),
        )}px`;
        fitRenderedPages(
          frame.viewport,
          frame.sizer,
          frame.scaleLayer,
          zoomRef.current,
        );
        if (!isCurrent()) return;

        const previousFrame = committedFrameRef.current;
        previousFrame?.dispose();
        shadowRoot.replaceChildren(
          previewStyleElement,
          frame.styleContainer,
          frame.viewport,
        );
        measurementHost = null;

        let active = true;
        let animationFrame = 0;
        let resizeObserver: ResizeObserver | undefined;
        const navigation = createReviewNavigation(
          frame.renderTarget,
          host.closest<HTMLElement>(".preview-viewport"),
          (entityId) => onSelectEntityRef.current?.(entityId),
        );
        const controller = {
          file,
          pageCount,
          navigation,
          fit: () => {
            if (!active || committedFrameRef.current !== controller) {
              return;
            }
            fitRenderedPages(
              frame.viewport,
              frame.sizer,
              frame.scaleLayer,
              zoomRef.current,
              widthCallback.current,
            );
          },
          dispose: () => {
            active = false;
            stopLiveUpdates();
            navigation.dispose();
            resizeObserver?.disconnect();
            cancelAnimationFrame(animationFrame);
            if (fitRef.current === controller.fit) {
              fitRef.current = null;
            }
          },
        };
        const scheduleFit = () => {
          if (!active || committedFrameRef.current !== controller) {
            return;
          }
          cancelAnimationFrame(animationFrame);
          animationFrame = requestAnimationFrame(controller.fit);
        };

        committedFrameRef.current = controller;
        fitRef.current = controller.fit;
        navigation.setSelection(selectionRef.current);
        resizeObserver = new ResizeObserver(scheduleFit);
        resizeObserver.observe(frame.viewport);
        resizeObserver.observe(frame.scaleLayer);
        controller.fit();

        void document.fonts?.ready.then(() => {
          if (active && committedFrameRef.current === controller) {
            scheduleFit();
          }
        });

        host.setAttribute("aria-busy", "false");
        onEntitiesChangeRef.current?.(initialLiveSnapshot?.entities ?? []);
        onLoadRef.current?.(pageCount);

        if (
          source &&
          incrementalLayout &&
          initialLiveSnapshot &&
          liveParagraphs
        ) {
          let appliedSnapshot = initialLiveSnapshot;
          let pendingSnapshot = initialLiveSnapshot;
          let pendingParagraphIds = new Set<string>();
          let running = false;
          let requestVersion = 0;
          let disposed = false;
          let prepared: PreparedLiveDocxUpdate | undefined;

          const enqueue = (
            snapshot: typeof initialLiveSnapshot,
            changedParagraphIds: Iterable<string>,
          ) => {
            if (disposed || snapshot.revision <= appliedSnapshot.revision) return;
            pendingSnapshot = snapshot;
            for (const id of changedParagraphIds) pendingParagraphIds.add(id);
            requestVersion += 1;
            if (!running) void flush();
          };

          const flush = async () => {
            if (running || disposed) return;
            running = true;
            try {
              while (!disposed && pendingParagraphIds.size) {
                const snapshot = pendingSnapshot;
                const paragraphIds = new Set(pendingParagraphIds);
                pendingParagraphIds.clear();
                const version = requestVersion;
                const updateIsCurrent = () =>
                  !disposed &&
                  active &&
                  committedFrameRef.current === controller &&
                  version === requestVersion;
                host.setAttribute("aria-busy", "true");
                try {
                  prepared = await incrementalLayout.prepare(
                    snapshot,
                    paragraphIds,
                    shadowRoot,
                    updateIsCurrent,
                    (container) => layoutDocxTabs(container, typography),
                  );
                  if (!updateIsCurrent()) {
                    prepared.discard();
                    prepared = undefined;
                    paragraphIds.forEach((id) => pendingParagraphIds.add(id));
                    continue;
                  }
                  const scrollViewport = host.closest<HTMLElement>(
                    ".preview-viewport",
                  );
                  const readingPosition = scrollViewport
                    ? captureReadingPosition(scrollViewport)
                    : undefined;
                  controller.pageCount = prepared.commit();
                  prepared = undefined;
                  navigation.refresh();
                  updatePageFields(frame.renderTarget);
                  layoutDocxTabs(frame.renderTarget, typography);
                  controller.fit();
                  appliedSnapshot = snapshot;
                  host.setAttribute("aria-busy", "false");
                  onLoadRef.current?.(controller.pageCount);
                  if (scrollViewport && readingPosition) {
                    requestAnimationFrame(() => {
                      if (
                        active &&
                        committedFrameRef.current === controller
                      ) {
                        restoreReadingPosition(
                          scrollViewport,
                          readingPosition,
                          true,
                        );
                      }
                    });
                  }
                } catch (error) {
                  prepared?.discard();
                  prepared = undefined;
                  paragraphIds.forEach((id) => pendingParagraphIds.add(id));
                  if (!updateIsCurrent()) continue;
                  pendingParagraphIds.clear();
                  host.setAttribute("aria-busy", "false");
                  onRefreshErrorRef.current?.(toErrorMessage(error));
                }
              }
            } finally {
              running = false;
            }
          };

          const receive = (change: LiveDocxChange) => {
            onEntitiesChangeRef.current?.(change.snapshot.entities);
            enqueue(change.snapshot, change.changedParagraphIds);
          };
          const unsubscribe = source.subscribe(receive);
          const latest = source.getSnapshot();
          if (latest.revision > appliedSnapshot.revision) {
            onEntitiesChangeRef.current?.(latest.entities);
          }
          const catchUpParagraphIds = liveCatchUpParagraphIds(
            appliedSnapshot,
            latest,
            liveParagraphs,
          );
          if (catchUpParagraphIds.size) enqueue(latest, catchUpParagraphIds);
          stopLiveUpdates = () => {
            if (disposed) return;
            disposed = true;
            requestVersion += 1;
            prepared?.discard();
            prepared = undefined;
            unsubscribe();
          };
        }
      } catch (error) {
        if (cancelled || renderSequence !== renderSequenceRef.current) {
          return;
        }

        measurementHost?.remove();
        measurementHost = null;
        host.setAttribute("aria-busy", "false");
        const committedFrame =
          committedFrameRef.current?.file === file
            ? committedFrameRef.current
            : null;
        if (committedFrame) {
          onLoadRef.current?.(committedFrame.pageCount);
          onRefreshErrorRef.current?.(toErrorMessage(error));
        } else {
          shadowRoot.replaceChildren(previewStyleElement);
          onErrorRef.current?.(toErrorMessage(error));
        }
      }
    };

    void render();

    return () => {
      cancelled = true;
      stopLiveUpdates();
      measurementHost?.remove();
    };
  }, [file, revisionMode, source, sourceError, sourcePending]);

  return (
    <div
      ref={hostRef}
      className="docx-preview"
      role="document"
      aria-label={`Preview of ${file.name}`}
      aria-busy="true"
    />
  );
}
