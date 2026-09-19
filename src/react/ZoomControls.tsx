import { useEffect, useState } from "react";

export function ZoomControls({
  zoom,
  fit,
  minimum,
  maximum,
  onChange,
}: {
  zoom: number;
  fit: boolean;
  minimum: number;
  maximum: number;
  onChange: (zoom: number | null) => void;
}) {
  const [draft, setDraft] = useState(String(zoom));
  useEffect(() => setDraft(String(zoom)), [zoom]);

  function draftZoom() {
    const value = Number(draft);
    return draft.trim() && Number.isFinite(value)
      ? Math.max(minimum, Math.min(maximum, Math.round(value)))
      : zoom;
  }

  function commit() {
    const next = draftZoom();
    setDraft(String(next));
    if (next !== zoom || fit) onChange(next);
  }

  function step(delta: number) {
    const next = Math.max(minimum, Math.min(maximum, draftZoom() + delta));
    setDraft(String(next));
    onChange(next);
  }

  return (
    <div className="docx-preview-zoom-controls" role="group" aria-label="Document zoom">
      <button
        type="button"
        aria-label="Zoom out"
        disabled={draftZoom() <= minimum}
        onClick={() => step(-10)}
      >
        −
      </button>
      <label className="docx-preview-zoom-value">
        <input
          aria-label="Zoom percentage"
          type="number"
          min={minimum}
          max={maximum}
          step={1}
          value={draft}
          onChange={(event) => setDraft(event.currentTarget.value)}
          onBlur={commit}
          onKeyDown={(event) => {
            if (event.key === "Enter") event.currentTarget.blur();
          }}
        />
        <span>%</span>
      </label>
      <button
        type="button"
        aria-label="Zoom in"
        disabled={draftZoom() >= maximum}
        onClick={() => step(10)}
      >
        +
      </button>
      <button
        type="button"
        className="docx-preview-zoom-fit"
        aria-pressed={fit}
        onClick={() => onChange(null)}
      >
        Fit width
      </button>
    </div>
  );
}
