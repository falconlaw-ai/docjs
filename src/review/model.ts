export type ReviewAnchor = {
  quote: string;
  prefix?: string;
  suffix?: string;
  paragraphId?: string;
};

export type ReviewItem = {
  id: string;
  kind: "comment" | "redline";
  anchor: ReviewAnchor;
  body?: string;
  replacement?: string;
  metadata?: unknown;
};

export type DocxSegment = { textId: string } | { literal: "\t" | "\n" };

export type DocxParagraph = {
  id: string;
  segments: DocxSegment[];
  /** Native w14:paraId, when present. It is only a hint within this file. */
  nativeId?: string;
  editable?: false;
  reason?: string;
};

export type DocxManifest = {
  schemaVersion: 1;
  paragraphs: DocxParagraph[];
  texts: Array<{ id: string; index: number }>;
};

export type ProjectedEntityStatus = "open" | "invalid" | "conflict";

export type ProjectedEntity = {
  id: string;
  kind: "comment" | "redline";
  body?: string;
  quote: string;
  replacement?: string;
  metadata?: unknown;
  status: ProjectedEntityStatus;
  reason?: string;
};

export type LiveDocxAnnotation = {
  entityId: string;
  kind: "comment" | "redline";
  start: number;
  end: number;
  first: boolean;
  last: boolean;
};

export type LiveDocxText = {
  id: string;
  value: string;
  annotations: readonly LiveDocxAnnotation[];
};

export type LiveDocxParagraph = {
  id: string;
  texts: readonly LiveDocxText[];
};

export type LiveDocxSnapshot = {
  revision: number;
  projectionMs: number;
  paragraphs: ReadonlyMap<string, LiveDocxParagraph>;
  entities: readonly ProjectedEntity[];
};

export type LiveDocxChange = {
  snapshot: LiveDocxSnapshot;
  changedParagraphIds: ReadonlySet<string>;
};

export type LiveDocxListener = (change: LiveDocxChange) => void;

export type LiveDocxSource = {
  readonly file: File;
  readonly manifest: DocxManifest;
  getSnapshot(): LiveDocxSnapshot;
  subscribe(listener: LiveDocxListener): () => void;
  dispose(): void;
};

export type ItemDocxSource = LiveDocxSource & {
  setItems(items: readonly ReviewItem[]): void;
};
