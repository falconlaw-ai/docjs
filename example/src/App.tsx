import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  DocumentPreview,
  type PreviewDocument,
  type PreviewMode,
  type ReviewItem,
  type WorkingDocumentPreparation,
} from "docx-preview/react";
import {
  fixtures,
  loadFixture,
  prepareFixture,
  type FixtureId,
} from "./fakeBackend";

const initialItems: readonly ReviewItem[] = [
  {
    id: "comment-services",
    kind: "comment",
    anchor: {
      quote: "Consultant will perform the services described in the Statement of Work",
    },
    body: "Confirm that the final statement of work is attached.",
    metadata: { severity: "medium" },
  },
  {
    id: "redline-consultant",
    kind: "redline",
    anchor: { quote: "Consultant is an independent contractor" },
    replacement: "Consultant acts as an independent professional",
    metadata: { severity: "high" },
  },
  {
    id: "redline-fees",
    kind: "redline",
    anchor: {
      quote: "Company will pay Consultant the fees specified in the Statement of Work",
    },
    replacement: "Company will pay the agreed fees",
    metadata: { severity: "low" },
  },
  {
    id: "invalid-anchor",
    kind: "comment",
    anchor: { quote: "This text is not present in the document." },
    body: "Invalid anchors remain visible as diagnostics.",
  },
];

const addedItem: ReviewItem = {
  id: "comment-payment",
  kind: "comment",
  anchor: { quote: "Company will pay the full amount of each such invoice" },
  body: "Check the invoice payment period.",
};

function fixtureFromUrl(): FixtureId {
  const value = new URL(window.location.href).searchParams.get("fixture");
  return fixtures.some((fixture) => fixture.id === value)
    ? (value as FixtureId)
    : "consulting-docx";
}

function modeFromUrl(): PreviewMode {
  return new URL(window.location.href).searchParams.get("initialMode") === "review"
    ? "review"
    : "original";
}

function formatForUpload(file: File): PreviewDocument["format"] | null {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "docx" || extension === "pdf") return extension;
  if (extension === "md" || extension === "markdown") return "md";
  return null;
}

async function digest(file: File) {
  const bytes = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return [...new Uint8Array(bytes)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

export function App() {
  const delay = Number(new URL(window.location.href).searchParams.get("delay") ?? 250);
  const customCards = new URL(window.location.href).searchParams.get("cards") === "custom";
  const inputRef = useRef<HTMLInputElement>(null);
  const loadGeneration = useRef(0);
  const preparationGeneration = useRef(0);
  const preparationController = useRef<AbortController | null>(null);
  const originalRef = useRef<PreviewDocument | null>(null);
  const [fixtureId, setFixtureId] = useState<FixtureId>(fixtureFromUrl);
  const [original, setOriginal] = useState<PreviewDocument | null>(null);
  const [mode, setMode] = useState<PreviewMode>(modeFromUrl);
  const [preparation, setPreparation] = useState<WorkingDocumentPreparation>({
    status: "unavailable",
  });
  const [items, setItems] = useState<readonly ReviewItem[]>(initialItems);
  const [itemVersion, setItemVersion] = useState(0);
  const [preparationCount, setPreparationCount] = useState(0);
  const [consumerAction, setConsumerAction] = useState("none");
  const [sourceDigest, setSourceDigest] = useState("pending");
  const [loadError, setLoadError] = useState<string | null>(null);

  useEffect(() => {
    const controller = new AbortController();
    const generation = ++loadGeneration.current;
    preparationGeneration.current += 1;
    preparationController.current?.abort();
    preparationController.current = null;
    setOriginal(null);
    originalRef.current = null;
    setPreparation({ status: "unavailable" });
    setItems(initialItems);
    setItemVersion(0);
    setLoadError(null);
    void loadFixture(fixtureId, controller.signal).then(
      (next) => {
        if (generation !== loadGeneration.current) return;
        originalRef.current = next;
        setOriginal(next);
      },
      (cause: unknown) => {
        if (controller.signal.aborted || generation !== loadGeneration.current) return;
        setLoadError(cause instanceof Error ? cause.message : "The fixture could not be loaded.");
      },
    );
    return () => controller.abort();
  }, [fixtureId]);

  useEffect(() => {
    if (!original) return;
    let active = true;
    setSourceDigest("pending");
    void digest(original.file).then((value) => {
      if (active) setSourceDigest(value);
    });
    return () => {
      active = false;
    };
  }, [original]);

  useEffect(
    () => () => {
      preparationGeneration.current += 1;
      preparationController.current?.abort();
    },
    [],
  );

  const requestPreparation = useCallback(
    (requested: PreviewDocument) => {
      preparationController.current?.abort();
      const controller = new AbortController();
      preparationController.current = controller;
      const generation = ++preparationGeneration.current;
      const identity = { id: requested.id, revision: requested.revision };
      setPreparationCount((count) => count + 1);
      setPreparation({ status: "pending", identity });
      void prepareFixture(requested, {
        delay: Number.isFinite(delay) ? Math.max(0, delay) : 250,
        signal: controller.signal,
      }).then(
        (file) => {
          const current = originalRef.current;
          if (
            controller.signal.aborted ||
            generation !== preparationGeneration.current ||
            current?.id !== requested.id ||
            current.revision !== requested.revision
          )
            return;
          setPreparation({ status: "ready", identity, file });
        },
        (cause: unknown) => {
          if (controller.signal.aborted || generation !== preparationGeneration.current) return;
          setPreparation({
            status: "error",
            identity,
            message: cause instanceof Error ? cause.message : "Preparation failed.",
          });
        },
      );
    },
    [delay],
  );

  const redlineCount = useMemo(
    () => items.filter((item) => item.kind === "redline").length,
    [items],
  );

  function chooseFixture(next: FixtureId) {
    const url = new URL(window.location.href);
    url.searchParams.set("fixture", next);
    window.history.replaceState(null, "", url);
    setMode("original");
    setFixtureId(next);
  }

  function openUpload(file: File) {
    const format = formatForUpload(file);
    if (!format) {
      setLoadError("Choose a PDF, DOCX, or Markdown file.");
      return;
    }
    loadGeneration.current += 1;
    preparationGeneration.current += 1;
    preparationController.current?.abort();
    const next: PreviewDocument = {
      id: `upload-${Date.now()}`,
      revision: `${file.lastModified}-${file.size}`,
      file,
      format,
    };
    originalRef.current = next;
    setOriginal(next);
    setMode("original");
    setPreparation({ status: "unavailable" });
    setItems(initialItems);
    setItemVersion(0);
    setLoadError(null);
  }

  function addItem() {
    setItems((current) =>
      current.some((item) => item.id === addedItem.id) ? current : [...current, addedItem],
    );
    setItemVersion((version) => version + 1);
  }

  function updateItem() {
    setItems((current) =>
      current.map((item) =>
        item.id === "redline-fees"
          ? {
              ...item,
              replacement:
                item.replacement === "Company will pay the agreed fees"
                  ? "Company will pay all agreed professional fees after receiving a valid invoice"
                  : "Company will pay the agreed fees",
            }
          : item,
      ),
    );
    setItemVersion((version) => version + 1);
  }

  function removeItem() {
    setItems((current) => current.filter((item) => item.id !== "redline-consultant"));
    setItemVersion((version) => version + 1);
  }

  return (
    <div className="example-app">
      <header className="example-toolbar">
        <div>
          <span className="example-eyebrow">Built package</span>
          <h1>Document preview example</h1>
        </div>
        <div className="example-actions">
          <label>
            <span>Example document</span>
            <select
              aria-label="Example document"
              value={fixtureId}
              onChange={(event) => chooseFixture(event.currentTarget.value as FixtureId)}
            >
              {fixtures.map((fixture) => (
                <option key={fixture.id} value={fixture.id}>
                  {fixture.label}
                </option>
              ))}
            </select>
          </label>
          <button type="button" onClick={() => inputRef.current?.click()}>
            Open a document
          </button>
          <input
            ref={inputRef}
            className="example-file-input"
            aria-label="Open a document"
            type="file"
            accept=".pdf,.docx,.md,.markdown"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file) openUpload(file);
              event.currentTarget.value = "";
            }}
          />
        </div>
      </header>
      <div className="example-test-controls" aria-label="Deterministic review controls">
        <button type="button" data-testid="item-add" onClick={addItem}>Add item</button>
        <button type="button" data-testid="item-update" onClick={updateItem}>Update item</button>
        <button type="button" data-testid="item-remove" onClick={removeItem}>Remove selected item</button>
        <span>Item version <output data-testid="item-version">{itemVersion}</output></span>
        <span>Preparation requests <output data-testid="preparation-count">{preparationCount}</output></span>
        <span>Redlines <output data-testid="supplied-redline-count">{redlineCount}</output></span>
        <span>Consumer action <output data-testid="consumer-action">{consumerAction}</output></span>
        <span className="example-digest">Source SHA-256 <output data-testid="source-digest">{sourceDigest}</output></span>
      </div>
      <div className="example-preview">
        {loadError ? (
          <div className="example-error" role="alert">{loadError}</div>
        ) : original ? (
          <DocumentPreview
            original={original}
            mode={mode}
            onModeChange={setMode}
            preparation={preparation}
            onRequestPreparation={requestPreparation}
            items={items}
            pdfAssets={{
              workerUrl: new URL("/pdfjs/pdf.worker.min.mjs", window.location.origin).href,
              resourceBaseUrl: new URL("/pdfjs/", window.location.origin).href,
            }}
            renderReviewItem={
              customCards
                ? ({ item, diagnostic }) => (
                    <div className="example-custom-card">
                      <strong>Consumer review card</strong>
                      <p>{item.anchor.quote}</p>
                      {item.replacement && <p>Replace with: {item.replacement}</p>}
                      {diagnostic && <p>{diagnostic}</p>}
                      <button
                        type="button"
                        data-review-action
                        onClick={() => setConsumerAction(item.id)}
                      >
                        Record consumer action
                      </button>
                    </div>
                  )
                : undefined
            }
          />
        ) : (
          <div className="example-loading" role="status">Loading fixture…</div>
        )}
      </div>
    </div>
  );
}
