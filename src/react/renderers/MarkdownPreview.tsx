import { useEffect, useRef, useState } from "react";
import ReactMarkdown, {
  defaultUrlTransform,
  type Components,
  type Options as ReactMarkdownOptions,
  type UrlTransform,
} from "react-markdown";
import rehypeExternalLinks from "rehype-external-links";
import rehypeSanitize, { defaultSchema } from "rehype-sanitize";
import remarkGfm from "remark-gfm";

import "./markdown-preview.css";

export interface MarkdownPreviewProps {
  file: File;
  zoom?: number;
  onPageWidth?: (width: number) => void;
  onLoad?: (pages: number) => void;
  onError?: (message: string) => void;
}

type PreviewState =
  | { status: "loading" }
  | { status: "ready"; source: string }
  | { status: "error"; message: string };

const EMBEDDED_IMAGE_PATTERN =
  /^data:image\/(?:avif|gif|jpeg|png|webp);base64,[a-z0-9+/=\s]+$/i;

const sanitizeSchema = {
  ...defaultSchema,
  attributes: {
    ...defaultSchema.attributes,
    a: [...(defaultSchema.attributes?.a ?? []), "target", "rel"],
  },
  protocols: {
    ...defaultSchema.protocols,
    src: [...(defaultSchema.protocols?.src ?? []), "data"],
  },
};

const rehypePlugins: NonNullable<ReactMarkdownOptions["rehypePlugins"]> = [
  [rehypeExternalLinks, { target: "_blank", rel: ["noreferrer", "noopener"] }],
  [rehypeSanitize, sanitizeSchema],
];

const transformUrl: UrlTransform = (url, key, node) => {
  if (node.tagName === "img" && key === "src") {
    return EMBEDDED_IMAGE_PATTERN.test(url) ? url : "";
  }

  return defaultUrlTransform(url);
};

const components: Components = {
  a: ({ node: _node, ...props }) => <a {...props} />,
  img: ({ node: _node, alt = "", src, ...props }) => {
    if (typeof src !== "string" || !EMBEDDED_IMAGE_PATTERN.test(src)) {
      return (
        <span className="markdown-preview__image-placeholder" role="note">
          {alt ? `[Image: ${alt}]` : "[Image not shown]"}
        </span>
      );
    }

    return (
      <img {...props} alt={alt} decoding="async" loading="lazy" src={src} />
    );
  },
  table: ({ node: _node, ...props }) => (
    <div className="markdown-preview__table-scroll">
      <table {...props} />
    </div>
  ),
};

export function MarkdownPreview({
  file,
  onLoad,
  onError,
  zoom = 100,
  onPageWidth,
}: MarkdownPreviewProps) {
  const [state, setState] = useState<PreviewState>({ status: "loading" });
  const onLoadRef = useRef(onLoad);
  const onErrorRef = useRef(onError);
  useEffect(() => {
    onPageWidth?.(794);
  }, [onPageWidth]);

  useEffect(() => {
    onLoadRef.current = onLoad;
    onErrorRef.current = onError;
  }, [onLoad, onError]);

  useEffect(() => {
    let active = true;

    setState({ status: "loading" });

    void file.text().then(
      (source) => {
        if (!active) return;

        setState({ status: "ready", source });
        onLoadRef.current?.(1);
      },
      () => {
        if (!active) return;

        const message = `Try choosing the file again. We couldn't read "${file.name}".`;
        setState({ status: "error", message });
        onErrorRef.current?.(message);
      },
    );

    return () => {
      active = false;
    };
  }, [file]);

  return (
    <article
      className="markdown-preview"
      style={{ width: 794, zoom: zoom / 100 }}
      aria-busy={state.status === "loading"}
    >
      {state.status === "loading" && (
        <p className="markdown-preview__status">Loading preview...</p>
      )}

      {state.status === "error" && (
        <p className="markdown-preview__status" role="alert">
          {state.message}
        </p>
      )}

      {state.status === "ready" && (
        <ReactMarkdown
          components={components}
          rehypePlugins={rehypePlugins}
          remarkPlugins={[remarkGfm]}
          skipHtml
          urlTransform={transformUrl}
        >
          {state.source}
        </ReactMarkdown>
      )}
    </article>
  );
}
