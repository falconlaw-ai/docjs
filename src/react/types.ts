import type { CSSProperties, ReactNode } from "react";
import type { ProjectedEntity, ReviewItem } from "../review/model";

export type { ProjectedEntity, ReviewItem } from "../review/model";

export type DocumentFormat = "docx" | "pdf" | "md";
export type PreviewMode = "original" | "review" | "final";
export type DocumentRevision = string | number;

export type PreviewDocument = {
  id: string;
  revision: DocumentRevision;
  format: DocumentFormat;
  file: File;
};

export type PreparationIdentity = Pick<PreviewDocument, "id" | "revision">;

export type WorkingDocumentPreparation =
  | { status: "unavailable" }
  | { status: "pending"; identity: PreparationIdentity }
  | {
      status: "ready";
      identity: PreparationIdentity;
      file: File;
    }
  | {
      status: "error";
      identity: PreparationIdentity;
      message: string;
    };

export type PdfPreviewAssets = {
  workerUrl: string;
  resourceBaseUrl: string;
};

export type ReviewItemRenderContext = {
  item: ReviewItem;
  entity: ProjectedEntity | null;
  selected: boolean;
  diagnostic: string | null;
};

export type DocumentPreviewProps = {
  original: PreviewDocument;
  mode: PreviewMode;
  onModeChange: (mode: PreviewMode) => void;
  preparation?: WorkingDocumentPreparation;
  onRequestPreparation?: (original: PreviewDocument) => void;
  items?: readonly ReviewItem[];
  renderReviewItem?: (context: ReviewItemRenderContext) => ReactNode;
  pdfAssets?: PdfPreviewAssets;
  className?: string;
  style?: CSSProperties;
  onLoad?: (result: { mode: PreviewMode; pages: number }) => void;
  onError?: (message: string) => void;
  onRefreshError?: (message: string) => void;
};
