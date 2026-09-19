import {
  Component,
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { PreviewModeControl } from "./PreviewModeControl";
import { ReviewPanel } from "./ReviewPanel";
import { ZoomControls } from "./ZoomControls";
import type { ReviewSelectionRequest } from "./renderers/reviewNavigation";
import {
  captureReadingPosition,
  restoreReadingPosition,
  type ReadingPosition,
} from "./renderers/readingPosition";
import type {
  DocumentPreviewProps,
  PdfPreviewAssets,
  PreviewDocument,
  PreviewMode,
  ProjectedEntity,
  ReviewItem,
  WorkingDocumentPreparation,
} from "./types";
import "./styles.css";

const DocxPreview = lazy(() =>
  import("./renderers/DocxPreview").then((module) => ({ default: module.DocxPreview })),
);
const PdfPreview = lazy(() =>
  import("./renderers/PdfPreview").then((module) => ({ default: module.PdfPreview })),
);
const MarkdownPreview = lazy(() =>
  import("./renderers/MarkdownPreview").then((module) => ({ default: module.MarkdownPreview })),
);

type PreviewView = {
  key: string;
  identityKey: string;
  mode: PreviewMode;
  format: PreviewDocument["format"];
  file: File;
};

type PendingReadingPosition = {
  identityKey: string;
  mode: PreviewMode;
  position: ReadingPosition;
};

type ZoomState = {
  minimum: number;
  maximum: number;
  choice: number | null;
};

type PendingReviewZoom = {
  identityKey: string;
  percentage: number;
};

const ABSOLUTE_MINIMUM_ZOOM = 20;
const ABSOLUTE_MAXIMUM_ZOOM = 200;

function clampZoom(zoom: number, minimum: number, maximum: number) {
  return Math.max(minimum, Math.min(maximum, zoom));
}

function viewportContentWidth(viewport: HTMLElement) {
  const computed = getComputedStyle(viewport);
  return Math.max(
    1,
    viewport.clientWidth -
      parseFloat(computed.paddingLeft) -
      parseFloat(computed.paddingRight),
  );
}

function normalizeZoomRange(minimum: number | undefined, maximum: number | undefined) {
  const effectiveMinimum = typeof minimum === "number" && Number.isFinite(minimum)
    ? clampZoom(minimum, ABSOLUTE_MINIMUM_ZOOM, ABSOLUTE_MAXIMUM_ZOOM)
    : ABSOLUTE_MINIMUM_ZOOM;
  const requestedMaximum = typeof maximum === "number" && Number.isFinite(maximum)
    ? clampZoom(maximum, ABSOLUTE_MINIMUM_ZOOM, ABSOLUTE_MAXIMUM_ZOOM)
    : ABSOLUTE_MAXIMUM_ZOOM;
  return {
    minimum: effectiveMinimum,
    maximum: Math.max(effectiveMinimum, requestedMaximum),
  };
}

function normalizeDefaultZoom(
  zoom: number | null | undefined,
  minimum: number,
  maximum: number,
) {
  return typeof zoom === "number" && Number.isFinite(zoom)
    ? clampZoom(zoom, minimum, maximum)
    : null;
}

const READING_NAVIGATION_KEYS = new Set([
  " ",
  "ArrowDown",
  "ArrowUp",
  "End",
  "Home",
  "PageDown",
  "PageUp",
]);

const fileIds = new WeakMap<File, number>();
let nextFileId = 0;

function fileId(file: File) {
  let id = fileIds.get(file);
  if (id === undefined) {
    id = ++nextFileId;
    fileIds.set(file, id);
  }
  return id;
}

function identityKey(original: Pick<PreviewDocument, "id" | "revision">) {
  return JSON.stringify([original.id, original.revision]);
}

function matchesOriginal(
  preparation: Exclude<WorkingDocumentPreparation, { status: "unavailable" }>,
  original: PreviewDocument,
) {
  return (
    preparation.identity.id === original.id &&
    preparation.identity.revision === original.revision
  );
}

function DocumentIcon() {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8l-5-5Z"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
      <path
        d="M14 3v5h5M8 12h8M8 16h6"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
      />
    </svg>
  );
}

class PreviewBoundary extends Component<
  { children: ReactNode; onError: (message: string) => void },
  { failed: boolean }
> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error) {
    this.props.onError(error.message || "The document could not be displayed.");
  }

  render() {
    return this.state.failed ? null : this.props.children;
  }
}

function MissingPdfAssets({ onError }: { onError: (message: string) => void }) {
  useEffect(() => {
    onError("PDF preview requires a worker URL and resource base URL.");
  }, [onError]);
  return null;
}

function ViewRenderer({
  view,
  items,
  zoom,
  pdfAssets,
  selection,
  onLoad,
  onError,
  onRefreshError,
  onPageWidth,
  onSelectEntity,
  onEntitiesChange,
}: {
  view: PreviewView;
  items: readonly ReviewItem[];
  zoom: number;
  pdfAssets?: PdfPreviewAssets;
  selection: ReviewSelectionRequest | null;
  onLoad: (view: PreviewView, pages: number) => void;
  onError: (view: PreviewView, message: string) => void;
  onRefreshError: (message: string) => void;
  onPageWidth: (view: PreviewView, width: number) => void;
  onSelectEntity: (entityId: string) => void;
  onEntitiesChange: (view: PreviewView, entities: readonly ProjectedEntity[]) => void;
}) {
  const loaded = useCallback((pages: number) => onLoad(view, pages), [onLoad, view]);
  const failed = useCallback((message: string) => onError(view, message), [onError, view]);
  const entitiesChanged = useCallback(
    (entities: readonly ProjectedEntity[]) => onEntitiesChange(view, entities),
    [onEntitiesChange, view],
  );
  const pageWidthChanged = useCallback(
    (width: number) => onPageWidth(view, width),
    [onPageWidth, view],
  );

  if (view.format === "pdf") {
    return pdfAssets ? (
      <PdfPreview
        file={view.file}
        zoom={zoom}
        assets={pdfAssets}
        onPageWidth={pageWidthChanged}
        onLoad={loaded}
        onError={failed}
      />
    ) : (
      <MissingPdfAssets onError={failed} />
    );
  }

  if (view.format === "md") {
    return (
      <MarkdownPreview
        file={view.file}
        zoom={zoom}
        onPageWidth={pageWidthChanged}
        onLoad={loaded}
        onError={failed}
      />
    );
  }

  const workingView = view.mode !== "original";
  return (
    <DocxPreview
      file={view.file}
      items={workingView ? items : undefined}
      revisionMode={view.mode === "final" ? "final" : "review"}
      zoom={zoom}
      onPageWidth={pageWidthChanged}
      onLoad={loaded}
      onError={failed}
      onRefreshError={onRefreshError}
      selection={workingView ? selection : null}
      onSelectEntity={workingView ? onSelectEntity : undefined}
      onEntitiesChange={workingView ? entitiesChanged : undefined}
    />
  );
}

export function DocumentPreview({
  original,
  mode,
  onModeChange,
  preparation = { status: "unavailable" },
  onRequestPreparation,
  items = [],
  renderReviewItem,
  pdfAssets,
  minZoom,
  maxZoom,
  defaultZoom,
  showModeControl = true,
  showZoomControls = true,
  showReviewPanel = true,
  className,
  style,
  onLoad,
  onError,
  onRefreshError,
}: DocumentPreviewProps) {
  const viewportRef = useRef<HTMLElement>(null);
  const selectionSequence = useRef(0);
  const requestedPreparationKey = useRef<string | null>(null);
  const readingPositions = useRef(
    new Map<string, Map<PreviewMode, ReadingPosition>>(),
  );
  const pendingPosition = useRef<PendingReadingPosition | null>(null);
  const panelRestorePosition = useRef<PendingReadingPosition | null>(null);
  const restoreGeneration = useRef(0);
  const previewReady = useRef(false);
  const committedViewRef = useRef<PreviewView | null>(null);
  const previousDesiredKey = useRef<string | null>(null);
  const requestedModeRef = useRef(mode);
  requestedModeRef.current = mode;
  const requestedZoomRange = normalizeZoomRange(minZoom, maxZoom);
  const [zoomState, setZoomState] = useState<ZoomState>(() => ({
    ...requestedZoomRange,
    choice: normalizeDefaultZoom(
      defaultZoom,
      requestedZoomRange.minimum,
      requestedZoomRange.maximum,
    ),
  }));
  const [pageWidth, setPageWidth] = useState(794);
  const [viewportWidth, setViewportWidth] = useState(794);
  const [committed, setCommitted] = useState<PreviewView | null>(null);
  const [pages, setPages] = useState<number | null>(null);
  const [renderError, setRenderError] = useState<{ key: string; message: string } | null>(null);
  const [refreshError, setRefreshError] = useState<string | null>(null);
  const [entities, setEntities] = useState<readonly ProjectedEntity[]>([]);
  const [projectionReady, setProjectionReady] = useState(false);
  const [selection, setSelection] = useState<ReviewSelectionRequest | null>(null);
  const zoomChoiceRef = useRef(zoomState.choice);
  zoomChoiceRef.current = zoomState.choice;

  const originalIdentityKey = identityKey(original);
  const previousMode = useRef({ identityKey: originalIdentityKey, mode });
  const pendingReviewZoom = useRef<PendingReviewZoom | null>(null);
  const pageWidths = useRef(new Map<string, number>());
  const previousZoomProps = useRef({
    identityKey: originalIdentityKey,
    defaultZoom,
    minimum: requestedZoomRange.minimum,
    maximum: requestedZoomRange.maximum,
  });
  const matchingPreparation =
    preparation.status !== "unavailable" && matchesOriginal(preparation, original)
      ? preparation
      : null;
  const workingFile =
    matchingPreparation?.status === "ready"
      ? matchingPreparation.file
      : original.format === "docx"
        ? original.file
        : null;
  const zoom =
    zoomState.choice ??
    clampZoom(
      Math.floor((viewportWidth / pageWidth) * 100),
      zoomState.minimum,
      zoomState.maximum,
    );

  const desired = useMemo<PreviewView | null>(() => {
    if (mode === "original" || !workingFile) {
      return {
        key: `${originalIdentityKey}:original:${fileId(original.file)}`,
        identityKey: originalIdentityKey,
        mode: "original",
        format: original.format,
        file: original.file,
      };
    }
    return {
      key: `${originalIdentityKey}:${mode}:${fileId(workingFile)}`,
      identityKey: originalIdentityKey,
      mode,
      format: "docx",
      file: workingFile,
    };
  }, [mode, original.file, original.format, originalIdentityKey, workingFile]);

  const desiredKeyRef = useRef(desired?.key ?? null);
  desiredKeyRef.current = desired?.key ?? null;
  const validCommitted = committed?.identityKey === originalIdentityKey ? committed : null;
  committedViewRef.current = validCommitted;
  const target = desired?.key === validCommitted?.key ? null : desired;
  const visible = validCommitted ?? target;
  const isStaging = Boolean(validCommitted && target);

  const cancelPendingRestore = useCallback(() => {
    pendingPosition.current = null;
    restoreGeneration.current += 1;
  }, []);

  const captureForView = useCallback((view: PreviewView) => {
    const viewport = viewportRef.current;
    if (!viewport) return null;
    const position = captureReadingPosition(viewport);
    let positionsByMode = readingPositions.current.get(view.identityKey);
    if (!positionsByMode) {
      positionsByMode = new Map();
      readingPositions.current.set(view.identityKey, positionsByMode);
    }
    positionsByMode.set(view.mode, position);
    return position;
  }, []);

  const restorePendingPosition = useCallback((view: PreviewView) => {
    const pending = pendingPosition.current;
    if (
      !pending ||
      pending.identityKey !== view.identityKey ||
      pending.mode !== view.mode
    ) {
      return;
    }
    const generation = ++restoreGeneration.current;
    requestAnimationFrame(() =>
      requestAnimationFrame(() => {
        const viewport = viewportRef.current;
        if (
          !viewport ||
          generation !== restoreGeneration.current ||
          pendingPosition.current !== pending
        ) {
          return;
        }
        restoreReadingPosition(viewport, pending.position);
        window.setTimeout(() => {
          if (
            generation === restoreGeneration.current &&
            pendingPosition.current === pending &&
            viewport === viewportRef.current
          ) {
            restoreReadingPosition(viewport, pending.position, true);
            pendingPosition.current = null;
          }
        }, 200);
      }),
    );
  }, []);

  useLayoutEffect(() => {
    const panelPending = panelRestorePosition.current;
    panelRestorePosition.current = null;
    const current = committedViewRef.current;
    if (
      current &&
      panelPending &&
      pendingPosition.current === panelPending
    ) {
      restorePendingPosition(current);
    }
    return () => {
      panelRestorePosition.current = null;
      const view = committedViewRef.current;
      if (
        !view ||
        requestedModeRef.current !== "review" ||
        zoomChoiceRef.current !== null ||
        !previewReady.current ||
        desiredKeyRef.current !== view.key ||
        pendingPosition.current
      ) {
        return;
      }
      const position = captureForView(view);
      if (position) {
        const pending = {
          identityKey: view.identityKey,
          mode: view.mode,
          position,
        };
        pendingPosition.current = pending;
        panelRestorePosition.current = pending;
      }
    };
  }, [captureForView, restorePendingPosition, showReviewPanel]);

  useLayoutEffect(() => {
    const nextKey = desired?.key ?? null;
    if (nextKey === previousDesiredKey.current) return;
    previousDesiredKey.current = nextKey;
    restoreGeneration.current += 1;

    if (!desired) {
      pendingPosition.current = null;
      previewReady.current = false;
      return;
    }

    const current = committedViewRef.current;
    if (current?.key === desired.key) {
      const saved = readingPositions.current
        .get(desired.identityKey)
        ?.get(desired.mode);
      pendingPosition.current = saved
        ? {
            identityKey: desired.identityKey,
            mode: desired.mode,
            position: saved,
          }
        : null;
      previewReady.current = true;
      if (saved) restorePendingPosition(current);
      return;
    }

    let currentPosition =
      pendingPosition.current?.identityKey === desired.identityKey
        ? pendingPosition.current.position
        : null;
    if (
      current &&
      current.identityKey === desired.identityKey &&
      previewReady.current
    ) {
      currentPosition = captureForView(current);
    }

    const saved = readingPositions.current
      .get(desired.identityKey)
      ?.get(desired.mode);
    const carriesCurrentPosition =
      current?.identityKey === desired.identityKey &&
      current.format === "docx" &&
      (desired.mode === "review" || desired.mode === "final");
    const position = saved ?? (carriesCurrentPosition ? currentPosition : null);
    pendingPosition.current = position
      ? { identityKey: desired.identityKey, mode: desired.mode, position }
      : null;
    previewReady.current = false;
  }, [captureForView, desired, restorePendingPosition]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    let scrollFrame = 0;
    const cancel = () => cancelPendingRestore();
    const cancelFromKey = (event: globalThis.KeyboardEvent) => {
      if (READING_NAVIGATION_KEYS.has(event.key)) cancel();
    };
    const saveScrolledPosition = () => {
      cancelAnimationFrame(scrollFrame);
      scrollFrame = requestAnimationFrame(() => {
        const current = committedViewRef.current;
        if (current && previewReady.current && !pendingPosition.current) {
          captureForView(current);
        }
      });
    };
    viewport.addEventListener("wheel", cancel, { passive: true });
    viewport.addEventListener("touchstart", cancel, { passive: true });
    viewport.addEventListener("pointerdown", cancel);
    viewport.addEventListener("keydown", cancelFromKey);
    viewport.addEventListener("scroll", saveScrolledPosition, { passive: true });
    return () => {
      cancelAnimationFrame(scrollFrame);
      viewport.removeEventListener("wheel", cancel);
      viewport.removeEventListener("touchstart", cancel);
      viewport.removeEventListener("pointerdown", cancel);
      viewport.removeEventListener("keydown", cancelFromKey);
      viewport.removeEventListener("scroll", saveScrolledPosition);
    };
  }, [cancelPendingRestore, captureForView, originalIdentityKey]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const measure = () => {
      setViewportWidth(viewportContentWidth(viewport));
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, [originalIdentityKey]);

  useEffect(() => {
    if (workingFile || mode !== "review" || matchingPreparation?.status === "pending") return;
    if (matchingPreparation?.status === "error") return;
    if (requestedPreparationKey.current === originalIdentityKey) return;
    requestedPreparationKey.current = originalIdentityKey;
    onRequestPreparation?.(original);
  }, [matchingPreparation, mode, onRequestPreparation, original, originalIdentityKey, workingFile]);

  useEffect(() => {
    pageWidths.current.clear();
    pendingReviewZoom.current = null;
    setPageWidth(794);
    setPages(null);
    setRenderError(null);
    setRefreshError(null);
    setEntities([]);
    setProjectionReady(false);
    setSelection(null);
  }, [originalIdentityKey]);

  useEffect(() => {
    if (!selection) return;
    const itemExists = items.some((item) => item.id === selection.entityId);
    const entity = entities.find((candidate) => candidate.id === selection.entityId);
    if (!itemExists || (projectionReady && entity?.status !== "open")) setSelection(null);
  }, [entities, items, projectionReady, selection]);

  const commitView = useCallback(
    (view: PreviewView, pageCount: number) => {
      if (desiredKeyRef.current !== view.key) return;
      previewReady.current = true;
      setCommitted(view);
      setPages(pageCount);
      setRenderError(null);
      setRefreshError(null);
      onLoad?.({ mode: view.mode, pages: pageCount });
      restorePendingPosition(view);
    },
    [onLoad, restorePendingPosition],
  );

  const failView = useCallback(
    (view: PreviewView, message: string) => {
      if (desiredKeyRef.current !== view.key) return;
      previewReady.current = false;
      cancelPendingRestore();
      setRenderError({ key: view.key, message });
      onError?.(message);
    },
    [cancelPendingRestore, onError],
  );

  const refreshFailed = useCallback(
    (message: string) => {
      setRefreshError(message);
      onRefreshError?.(message);
    },
    [onRefreshError],
  );

  const entitiesChanged = useCallback(
    (view: PreviewView, next: readonly ProjectedEntity[]) => {
      if (view.identityKey !== identityKey(original)) return;
      setEntities(next);
      setProjectionReady(true);
    },
    [original],
  );

  const pageWidthChanged = useCallback((view: PreviewView, width: number) => {
    pageWidths.current.set(view.key, width);
    if (desiredKeyRef.current === view.key) setPageWidth(width);
  }, []);

  const selectEntity = useCallback(
    (entityId: string, origin: ReviewSelectionRequest["origin"]) => {
      cancelPendingRestore();
      setSelection({ entityId, origin, sequence: ++selectionSequence.current });
    },
    [cancelPendingRestore],
  );

  const applyZoomState = useCallback((next: ZoomState) => {
    const current = committedViewRef.current;
    if (
      current &&
      previewReady.current &&
      desiredKeyRef.current === current.key &&
      !pendingPosition.current
    ) {
      const position = captureForView(current);
      if (position) {
        pendingPosition.current = {
          identityKey: current.identityKey,
          mode: current.mode,
          position,
        };
      }
    }
    setZoomState(next);
    if (current && previewReady.current && desiredKeyRef.current === current.key) {
      restorePendingPosition(current);
    }
  }, [captureForView, restorePendingPosition]);

  function changeZoom(next: number | null) {
    pendingReviewZoom.current = null;
    const choice =
      next === null
        ? null
        : clampZoom(next, zoomState.minimum, zoomState.maximum);
    applyZoomState({ ...zoomState, choice });
  }

  useLayoutEffect(() => {
    const previous = previousMode.current;
    previousMode.current = { identityKey: originalIdentityKey, mode };
    if (previous.identityKey !== originalIdentityKey) {
      pendingReviewZoom.current = null;
      return;
    }
    if (previous.mode === mode) return;

    const previousProps = previousZoomProps.current;
    const zoomPropsChanged =
      !Object.is(previousProps.defaultZoom, defaultZoom) ||
      previousProps.minimum !== requestedZoomRange.minimum ||
      previousProps.maximum !== requestedZoomRange.maximum;
    if (zoomPropsChanged) {
      pendingReviewZoom.current = null;
      return;
    }

    if (mode === "review") {
      pendingReviewZoom.current = {
        identityKey: originalIdentityKey,
        percentage: zoom,
      };
      if (zoomState.choice === null) {
        applyZoomState({ ...zoomState, choice: zoom });
      }
      return;
    }

    pendingReviewZoom.current = null;
    if (previous.mode === "review" && !Object.is(zoomState.choice, zoom)) {
      applyZoomState({ ...zoomState, choice: zoom });
    }
  }, [
    applyZoomState,
    defaultZoom,
    mode,
    originalIdentityKey,
    requestedZoomRange.maximum,
    requestedZoomRange.minimum,
    zoom,
    zoomState,
  ]);

  useLayoutEffect(() => {
    const pending = pendingReviewZoom.current;
    if (
      !pending ||
      pending.identityKey !== originalIdentityKey ||
      mode !== "review" ||
      validCommitted?.mode !== "review" ||
      validCommitted.key !== desired?.key ||
      pages === null
    ) {
      return;
    }
    const workingPageWidth = pageWidths.current.get(validCommitted.key);
    const viewport = viewportRef.current;
    if (!workingPageWidth || !viewport) return;

    const availableWidth = viewportContentWidth(viewport);
    const availablePercentage = Math.floor((availableWidth / workingPageWidth) * 100);
    const choice = pending.percentage > availablePercentage
      ? null
      : clampZoom(pending.percentage, zoomState.minimum, zoomState.maximum);
    pendingReviewZoom.current = null;
    if (!Object.is(zoomState.choice, choice)) {
      applyZoomState({ ...zoomState, choice });
    }
  }, [
    applyZoomState,
    desired?.key,
    mode,
    originalIdentityKey,
    pageWidth,
    pages,
    validCommitted,
    zoomState,
  ]);

  useLayoutEffect(() => {
    const previous = previousZoomProps.current;
    const defaultChanged =
      previous.identityKey !== originalIdentityKey ||
      !Object.is(previous.defaultZoom, defaultZoom);
    const boundsChanged =
      previous.minimum !== requestedZoomRange.minimum ||
      previous.maximum !== requestedZoomRange.maximum;
    previousZoomProps.current = {
      identityKey: originalIdentityKey,
      defaultZoom,
      minimum: requestedZoomRange.minimum,
      maximum: requestedZoomRange.maximum,
    };
    if (!defaultChanged && !boundsChanged) return;
    pendingReviewZoom.current = null;

    const choice = defaultChanged
      ? normalizeDefaultZoom(
          defaultZoom,
          requestedZoomRange.minimum,
          requestedZoomRange.maximum,
        )
      : zoomState.choice === null
        ? null
        : clampZoom(
            zoomState.choice,
            requestedZoomRange.minimum,
            requestedZoomRange.maximum,
          );
    if (
      zoomState.minimum === requestedZoomRange.minimum &&
      zoomState.maximum === requestedZoomRange.maximum &&
      Object.is(zoomState.choice, choice)
    ) {
      return;
    }

    applyZoomState({ ...requestedZoomRange, choice });
  }, [
    applyZoomState,
    defaultZoom,
    originalIdentityKey,
    requestedZoomRange.maximum,
    requestedZoomRange.minimum,
    zoomState,
  ]);

  function changeMode(next: PreviewMode) {
    if (next === "final" && !workingFile) return;
    setRenderError(null);
    if (next === "review" && !workingFile && matchingPreparation?.status !== "pending") {
      requestedPreparationKey.current = originalIdentityKey;
      onRequestPreparation?.(original);
    }
    onModeChange(next);
  }

  const preparationPending =
    mode === "review" && !workingFile && matchingPreparation?.status === "pending";
  const preparationError =
    mode === "review" && !workingFile && matchingPreparation?.status === "error"
      ? matchingPreparation.message
      : null;
  const currentMatchesRequest = validCommitted?.key === desired?.key;
  const pageStatus = currentMatchesRequest && validCommitted && pages !== null
    ? `${pages} ${pages === 1 ? "page" : "pages"}`
    : null;
  const previewUpdate = preparationPending
    ? "Preparing review…"
    : target || !validCommitted
      ? "Preparing preview…"
      : "";
  const slots = [validCommitted, target].filter(
    (view, index, list): view is PreviewView =>
      Boolean(view) && list.findIndex((candidate) => candidate?.key === view?.key) === index,
  );
  const itemCountLabel = items.length === 1 ? "1 review item" : `${items.length} review items`;

  return (
    <section
      className={`docx-preview-shell${className ? ` ${className}` : ""}`}
      style={style}
      data-preview-mode={mode}
      data-committed-mode={currentMatchesRequest && pages !== null ? validCommitted?.mode : undefined}
    >
      <div className="docx-preview-bar">
        <div className="docx-preview-document-info">
          <DocumentIcon />
          <span className="docx-preview-filename" title={original.file.name}>
            {original.file.name}
          </span>
          <span className="docx-preview-format-badge">{original.format === "md" ? "Markdown" : original.format.toUpperCase()}</span>
          {pageStatus && (
            <span className="docx-preview-status" role="status" aria-label="Preview status" aria-live="polite">
              {pageStatus}
            </span>
          )}
        </div>
        {showModeControl && (
          <PreviewModeControl
            value={mode}
            finalDisabled={!workingFile}
            reviewPending={preparationPending}
            onChange={changeMode}
          />
        )}
        <div className="docx-preview-controls">
          {showZoomControls && (
            <ZoomControls
              zoom={zoom}
              fit={zoomState.choice === null}
              minimum={zoomState.minimum}
              maximum={zoomState.maximum}
              onChange={changeZoom}
            />
          )}
        </div>
      </div>
      <span className="docx-preview-announcement" role="status" aria-label="Preview update" aria-live="polite">
        {previewUpdate}
      </span>
      {preparationError && (
        <div className="docx-preview-alert" role="alert">
          <span>{preparationError}</span>
          {onRequestPreparation && (
            <button type="button" onClick={() => onRequestPreparation(original)}>
              Retry preparation
            </button>
          )}
        </div>
      )}
      {refreshError && (
        <div className="docx-preview-alert" role="alert">
          Couldn't refresh the preview. Showing the last complete version. {refreshError}
        </div>
      )}
      {renderError && renderError.key === desired?.key && (
        <div className="docx-preview-alert" role="alert">
          Couldn't preview this document. {renderError.message}
        </div>
      )}
      <div className={`docx-preview-layout${mode === "review" && showReviewPanel ? " docx-preview-layout--review" : ""}`}>
        <main
          ref={viewportRef}
          className="docx-preview-viewport preview-viewport"
          aria-label="Document preview"
        >
          {slots.map((view) => {
            const staged = isStaging && view.key === target?.key;
            return (
              <div
                key={view.key}
                className={`docx-preview-render-slot${staged ? " is-staged" : ""}`}
                aria-hidden={staged ? "true" : undefined}
              >
                <PreviewBoundary onError={(message) => failView(view, message)}>
                  <Suspense fallback={<div className="docx-preview-loading">Loading renderer…</div>}>
                    <ViewRenderer
                      view={view}
                      items={items}
                      zoom={zoom}
                      pdfAssets={pdfAssets}
                      selection={view.mode === "review" ? selection : null}
                      onLoad={commitView}
                      onError={failView}
                      onRefreshError={refreshFailed}
                      onPageWidth={pageWidthChanged}
                      onSelectEntity={(entityId) => selectEntity(entityId, "document")}
                      onEntitiesChange={entitiesChanged}
                    />
                  </Suspense>
                </PreviewBoundary>
              </div>
            );
          })}
          {!visible && !renderError && (
            <div className="docx-preview-loading">Preparing preview…</div>
          )}
        </main>
        {mode === "review" && showReviewPanel && (
          <ReviewPanel
            items={items}
            entities={entities}
            projectionReady={projectionReady}
            selection={selection}
            navigationEnabled={validCommitted?.mode === "review" && currentMatchesRequest}
            renderItem={renderReviewItem}
            onSelectEntity={(entityId) => selectEntity(entityId, "review")}
          />
        )}
      </div>
      <span className="docx-preview-item-count" aria-label={itemCountLabel}>
        {itemCountLabel}
      </span>
    </section>
  );
}
