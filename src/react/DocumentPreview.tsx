import {
  Component,
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { PreviewModeControl } from "./PreviewModeControl";
import { ReviewPanel } from "./ReviewPanel";
import { ZoomControls } from "./ZoomControls";
import type { ReviewSelectionRequest } from "./renderers/reviewNavigation";
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
  onPageWidth: (width: number) => void;
  onSelectEntity: (entityId: string) => void;
  onEntitiesChange: (view: PreviewView, entities: readonly ProjectedEntity[]) => void;
}) {
  const loaded = useCallback((pages: number) => onLoad(view, pages), [onLoad, view]);
  const failed = useCallback((message: string) => onError(view, message), [onError, view]);
  const entitiesChanged = useCallback(
    (entities: readonly ProjectedEntity[]) => onEntitiesChange(view, entities),
    [onEntitiesChange, view],
  );

  if (view.format === "pdf") {
    return pdfAssets ? (
      <PdfPreview
        file={view.file}
        zoom={zoom}
        assets={pdfAssets}
        onPageWidth={onPageWidth}
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
        onPageWidth={onPageWidth}
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
      onPageWidth={onPageWidth}
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
  className,
  style,
  onLoad,
  onError,
  onRefreshError,
}: DocumentPreviewProps) {
  const viewportRef = useRef<HTMLElement>(null);
  const selectionSequence = useRef(0);
  const requestedPreparationKey = useRef<string | null>(null);
  const [zoomChoice, setZoomChoice] = useState<number | null>(null);
  const [pageWidth, setPageWidth] = useState(794);
  const [viewportWidth, setViewportWidth] = useState(794);
  const [committed, setCommitted] = useState<PreviewView | null>(null);
  const [pages, setPages] = useState<number | null>(null);
  const [renderError, setRenderError] = useState<{ key: string; message: string } | null>(null);
  const [refreshError, setRefreshError] = useState<string | null>(null);
  const [entities, setEntities] = useState<readonly ProjectedEntity[]>([]);
  const [projectionReady, setProjectionReady] = useState(false);
  const [selection, setSelection] = useState<ReviewSelectionRequest | null>(null);

  const originalIdentityKey = identityKey(original);
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
    zoomChoice ??
    Math.max(25, Math.min(300, Math.floor((viewportWidth / pageWidth) * 100)));

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
  const target = desired?.key === validCommitted?.key ? null : desired;
  const visible = validCommitted ?? target;
  const isStaging = Boolean(validCommitted && target);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const measure = () => {
      const computed = getComputedStyle(viewport);
      setViewportWidth(
        Math.max(
          1,
          viewport.clientWidth -
            parseFloat(computed.paddingLeft) -
            parseFloat(computed.paddingRight),
        ),
      );
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
    setZoomChoice(null);
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
      setCommitted(view);
      setPages(pageCount);
      setRenderError(null);
      setRefreshError(null);
      onLoad?.({ mode: view.mode, pages: pageCount });
    },
    [onLoad],
  );

  const failView = useCallback(
    (view: PreviewView, message: string) => {
      if (desiredKeyRef.current !== view.key) return;
      setRenderError({ key: view.key, message });
      onError?.(message);
    },
    [onError],
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

  const selectEntity = useCallback(
    (entityId: string, origin: ReviewSelectionRequest["origin"]) => {
      setSelection({ entityId, origin, sequence: ++selectionSequence.current });
    },
    [],
  );

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
  const status = preparationPending
    ? "Preparing review…"
    : currentMatchesRequest && validCommitted
      ? `${pages ?? "Preparing"} ${pages === 1 ? "page" : "pages"} · ${validCommitted.mode === "original" ? "Original" : validCommitted.mode === "review" ? "Review" : "Final"} view · Read only`
      : target
        ? "Preparing preview…"
        : validCommitted
          ? `${validCommitted.mode === "original" ? "Original" : validCommitted.mode === "review" ? "Review" : "Final"} view · Read only`
          : "Preparing preview…";
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
    >
      <div className="docx-preview-bar">
        <div className="docx-preview-document-info">
          <DocumentIcon />
          <span className="docx-preview-filename" title={original.file.name}>
            {original.file.name}
          </span>
          <span className="docx-preview-format-badge">{original.format === "md" ? "Markdown" : original.format.toUpperCase()}</span>
        </div>
        <PreviewModeControl
          value={mode}
          finalDisabled={!workingFile}
          reviewPending={preparationPending}
          onChange={changeMode}
        />
        <div className="docx-preview-controls">
          <ZoomControls
            zoom={zoom}
            fit={zoomChoice === null}
            onChange={setZoomChoice}
          />
          <div className="docx-preview-status" role="status" aria-label="Preview status" aria-live="polite">
            {status}
          </div>
        </div>
      </div>
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
      <div className={`docx-preview-layout${mode === "review" ? " docx-preview-layout--review" : ""}`}>
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
                      onPageWidth={setPageWidth}
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
        {mode === "review" && (
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
