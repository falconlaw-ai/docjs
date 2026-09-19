import type { PreviewMode } from "./types";

type PreviewModeControlProps = {
  value: PreviewMode;
  finalDisabled: boolean;
  reviewPending?: boolean;
  onChange: (mode: PreviewMode) => void;
};

function OriginalIcon() {
  return (
    <svg viewBox="0 0 16 16" aria-hidden="true">
      <path d="M4.5 1.75h4.25l2.75 2.8v9.7h-7z" />
      <path d="M8.75 1.75v2.8h2.75M6.25 7.25h3.5M6.25 9.5h3.5" />
    </svg>
  );
}

function ReviewIcon() {
  return (
    <svg viewBox="0 0 16 16" aria-hidden="true">
      <path d="M3 3.25h6.5M3 6.5h4.25M3 9.75h3" />
      <path d="m7.25 11.75 1.8-.4 3.8-3.8-1.4-1.4-3.8 3.8zM10.9 6.7l1.4 1.4" />
    </svg>
  );
}

function FinalIcon() {
  return (
    <svg viewBox="0 0 16 16" aria-hidden="true">
      <path d="M3.25 8.25 6.5 11.5l6.25-7" />
    </svg>
  );
}

const options = [
  { value: "original" as const, label: "Original", icon: <OriginalIcon /> },
  { value: "review" as const, label: "Review", icon: <ReviewIcon /> },
  { value: "final" as const, label: "Final", icon: <FinalIcon /> },
];

export function PreviewModeControl({
  value,
  finalDisabled,
  reviewPending = false,
  onChange,
}: PreviewModeControlProps) {
  return (
    <div className="docx-preview-mode-control" role="group" aria-label="Document view">
      {options.map((option) => (
        <button
          key={option.value}
          type="button"
          className="docx-preview-mode-option"
          aria-label={option.label}
          aria-pressed={value === option.value}
          aria-busy={option.value === "review" && reviewPending}
          disabled={option.value === "final" && finalDisabled}
          title={
            option.value === "final" && finalDisabled
              ? "Prepare a review document before opening the final view"
              : undefined
          }
          onClick={() => onChange(option.value)}
        >
          {option.icon}
          <span>{option.label}</span>
        </button>
      ))}
    </div>
  );
}
