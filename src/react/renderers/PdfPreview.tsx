import {
  GlobalWorkerOptions,
  TextLayer,
  getDocument,
  type PDFDocumentLoadingTask,
  type PDFDocumentProxy,
  type RenderTask,
} from "pdfjs-dist";
import { useCallback, useEffect, useRef, useState } from "react";

import "./pdf-preview.css";

export type PdfPreviewAssets = {
  workerUrl: string;
  resourceBaseUrl: string;
};

export type PdfPreviewProps = {
  file: File;
  assets: PdfPreviewAssets;
  zoom?: number;
  onPageWidth?: (width: number) => void;
  onLoad?: (pages: number) => void;
  onError?: (message: string) => void;
};

type PageDescriptor = {
  height: number;
  pageNumber: number;
  width: number;
};

type LoadedDocument = {
  file: File;
  pages: PageDescriptor[];
  pdf: PDFDocumentProxy;
};

const PDF_TO_CSS_UNITS = 96 / 72;

const assetDirectory = (baseUrl: string, name: string) =>
  new URL(
    `${name}/`,
    new URL(baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`, document.baseURI),
  ).href;

const errorMessage = (error: unknown) => {
  if (error instanceof Error && error.message.trim()) {
    return error.message;
  }

  return "We couldn't open this PDF.";
};

const isCancelled = (error: unknown) =>
  error instanceof Error &&
  (error.name === "AbortException" ||
    error.name === "RenderingCancelledException");

function PdfPage({
  descriptor,
  zoom,
  onError,
  pdf,
}: {
  descriptor: PageDescriptor;
  zoom: number;
  onError: (message: string) => void;
  pdf: PDFDocumentProxy;
}) {
  const pageElementRef = useRef<HTMLElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const textLayerElementRef = useRef<HTMLDivElement>(null);
  const activeRenderRef = useRef<RenderTask | null>(null);
  const [isVisible, setIsVisible] = useState(false);
  const [isRendered, setIsRendered] = useState(false);
  const [renderWidth, setRenderWidth] = useState(0);
  const [renderError, setRenderError] = useState<string | null>(null);

  useEffect(() => {
    const element = pageElementRef.current;
    if (!element) {
      return;
    }

    const IntersectionObserverConstructor = globalThis.IntersectionObserver as
      typeof IntersectionObserver | undefined;

    if (!IntersectionObserverConstructor) {
      setIsVisible(true);
      return;
    }

    const observer = new IntersectionObserverConstructor(
      ([entry]) => {
        if (entry.isIntersecting) {
          setIsVisible(true);
          observer.disconnect();
        }
      },
      { rootMargin: "800px 0px" },
    );

    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    const element = pageElementRef.current;
    if (!element) {
      return;
    }

    const updateWidth = () => {
      const nextWidth = Math.round(element.getBoundingClientRect().width);
      setRenderWidth((currentWidth) =>
        nextWidth > 0 && nextWidth !== currentWidth ? nextWidth : currentWidth,
      );
    };

    updateWidth();

    const ResizeObserverConstructor = globalThis.ResizeObserver as
      typeof ResizeObserver | undefined;

    if (!ResizeObserverConstructor) {
      window.addEventListener("resize", updateWidth);
      return () => window.removeEventListener("resize", updateWidth);
    }

    const observer = new ResizeObserverConstructor(updateWidth);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    if (!isVisible || renderWidth === 0) {
      return;
    }

    const canvas = canvasRef.current;
    const textLayerElement = textLayerElementRef.current;
    if (!canvas || !textLayerElement) {
      return;
    }

    let disposed = false;
    let renderTask: RenderTask | null = null;
    let textLayer: TextLayer | null = null;
    const previousRender = activeRenderRef.current;
    previousRender?.cancel();

    const render = async () => {
      try {
        await previousRender?.promise.catch(() => undefined);
        if (disposed) {
          return;
        }

        setIsRendered(false);
        setRenderError(null);
        const page = await pdf.getPage(descriptor.pageNumber);
        if (disposed) {
          return;
        }

        const unscaledViewport = page.getViewport({ scale: 1 });
        const viewport = page.getViewport({
          scale: renderWidth / unscaledViewport.width,
        });
        const outputScale = Math.min(window.devicePixelRatio || 1, 2);
        const context = canvas.getContext("2d", { alpha: false });

        if (!context) {
          throw new Error("This browser couldn't create a PDF canvas.");
        }

        canvas.width = Math.floor(viewport.width * outputScale);
        canvas.height = Math.floor(viewport.height * outputScale);
        canvas.style.width = `${viewport.width}px`;
        canvas.style.height = `${viewport.height}px`;

        textLayerElement.replaceChildren();
        textLayerElement.style.setProperty(
          "--scale-factor",
          String(viewport.scale),
        );
        textLayerElement.style.setProperty(
          "--user-unit",
          String(viewport.userUnit),
        );
        textLayerElement.style.setProperty(
          "--total-scale-factor",
          String(viewport.scale * viewport.userUnit),
        );

        renderTask = page.render({
          background: "#ffffff",
          canvas,
          transform:
            outputScale === 1
              ? undefined
              : [outputScale, 0, 0, outputScale, 0, 0],
          viewport,
        });
        activeRenderRef.current = renderTask;

        textLayer = new TextLayer({
          container: textLayerElement,
          textContentSource: page.streamTextContent({
            includeMarkedContent: true,
          }),
          viewport,
        });

        await Promise.all([renderTask.promise, textLayer.render()]);
        if (!disposed) {
          setIsRendered(true);
        }
      } catch (error) {
        if (disposed || isCancelled(error)) {
          return;
        }

        const message = errorMessage(error);
        setRenderError(message);
        onError(message);
      } finally {
        if (activeRenderRef.current === renderTask) {
          activeRenderRef.current = null;
        }
      }
    };

    void render();

    return () => {
      disposed = true;
      renderTask?.cancel();
      textLayer?.cancel();
    };
  }, [descriptor.pageNumber, isVisible, onError, pdf, renderWidth]);

  return (
    <section
      aria-label={`Page ${descriptor.pageNumber}`}
      className="pdf-preview__page"
      ref={pageElementRef}
      style={{
        aspectRatio: `${descriptor.width} / ${descriptor.height}`,
        width: `${(descriptor.width * zoom) / 100}px`,
      }}
    >
      <canvas
        aria-hidden="true"
        className="pdf-preview__canvas"
        ref={canvasRef}
      />
      <div className="textLayer" ref={textLayerElementRef} />
      {!isRendered && !renderError && (
        <span className="pdf-preview__page-placeholder">
          Page {descriptor.pageNumber}
        </span>
      )}
      {renderError && (
        <p className="pdf-preview__page-error" role="alert">
          This page couldn't be rendered.
        </p>
      )}
    </section>
  );
}

export function PdfPreview({
  file,
  assets,
  onLoad,
  onError,
  zoom = 100,
  onPageWidth,
}: PdfPreviewProps) {
  const [loadedDocument, setLoadedDocument] = useState<LoadedDocument | null>(
    null,
  );
  const [loadError, setLoadError] = useState<{
    file: File;
    message: string;
  } | null>(null);
  const onLoadRef = useRef(onLoad);
  const onErrorRef = useRef(onError);
  const widthCallback = useRef(onPageWidth);
  widthCallback.current = onPageWidth;

  useEffect(() => {
    onLoadRef.current = onLoad;
  }, [onLoad]);

  useEffect(() => {
    onErrorRef.current = onError;
  }, [onError]);

  const reportError = useCallback((message: string) => {
    onErrorRef.current?.(message);
  }, []);

  useEffect(() => {
    let disposed = false;
    let loadingTask: PDFDocumentLoadingTask | null = null;

    setLoadedDocument(null);
    setLoadError(null);

    const load = async () => {
      try {
        GlobalWorkerOptions.workerSrc = assets.workerUrl;
        const data = new Uint8Array(await file.arrayBuffer());
        if (disposed) {
          return;
        }

        loadingTask = getDocument({
          cMapPacked: true,
          cMapUrl: assetDirectory(assets.resourceBaseUrl, "cmaps"),
          iccUrl: assetDirectory(assets.resourceBaseUrl, "iccs"),
          data,
          standardFontDataUrl: assetDirectory(
            assets.resourceBaseUrl,
            "standard_fonts",
          ),
          wasmUrl: assetDirectory(assets.resourceBaseUrl, "wasm"),
        });

        const pdf = await loadingTask.promise;
        const pages = await Promise.all(
          Array.from({ length: pdf.numPages }, async (_, index) => {
            const pageNumber = index + 1;
            const page = await pdf.getPage(pageNumber);
            const viewport = page.getViewport({ scale: PDF_TO_CSS_UNITS });

            return {
              height: viewport.height,
              pageNumber,
              width: viewport.width,
            };
          }),
        );

        if (disposed) {
          return;
        }

        setLoadedDocument({ file, pages, pdf });
        widthCallback.current?.(Math.max(...pages.map((page) => page.width)));
        onLoadRef.current?.(pdf.numPages);
      } catch (error) {
        if (disposed || isCancelled(error)) {
          return;
        }

        const message = errorMessage(error);
        setLoadError({ file, message });
        onErrorRef.current?.(message);
      }
    };

    void load();

    return () => {
      disposed = true;
      void loadingTask?.destroy().catch(() => undefined);
    };
  }, [assets.resourceBaseUrl, assets.workerUrl, file]);

  const currentDocument = loadedDocument?.file === file ? loadedDocument : null;
  const currentError = loadError?.file === file ? loadError.message : null;

  return (
    <div
      aria-busy={!currentDocument && !currentError}
      className="pdf-preview"
      data-testid="pdf-preview"
    >
      {currentError ? (
        <div
          className="pdf-preview__message pdf-preview__message--error"
          role="alert"
        >
          <strong>We couldn't open this PDF.</strong>
          <span>{currentError}</span>
        </div>
      ) : currentDocument ? (
        currentDocument.pages.map((page) => (
          <PdfPage
            zoom={zoom}
            descriptor={page}
            key={page.pageNumber}
            onError={reportError}
            pdf={currentDocument.pdf}
          />
        ))
      ) : (
        <div className="pdf-preview__message" role="status">
          Loading PDF…
        </div>
      )}
    </div>
  );
}
