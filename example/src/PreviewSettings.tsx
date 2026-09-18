import { useEffect, useLayoutEffect, useRef, useState } from "react";

type PreviewSettingsProps = {
  minimumZoom: string;
  maximumZoom: string;
  defaultZoom: string;
  showModeControl: boolean;
  showZoomControls: boolean;
  showReviewPanel: boolean;
  onMinimumZoomChange: (value: string) => void;
  onMaximumZoomChange: (value: string) => void;
  onDefaultZoomChange: (value: string) => void;
  onShowModeControlChange: (value: boolean) => void;
  onShowZoomControlsChange: (value: boolean) => void;
  onShowReviewPanelChange: (value: boolean) => void;
};

function SettingsIcon() {
  return (
    <svg viewBox="0 0 20 20" aria-hidden="true">
      <path d="M8.9 2.4h2.2l.5 1.8c.4.1.8.3 1.2.5l1.7-.8 1.6 1.6-.8 1.7c.2.4.4.8.5 1.2l1.8.5v2.2l-1.8.5c-.1.4-.3.8-.5 1.2l.8 1.7-1.6 1.6-1.7-.8c-.4.2-.8.4-1.2.5l-.5 1.8H8.9l-.5-1.8c-.4-.1-.8-.3-1.2-.5l-1.7.8-1.6-1.6.8-1.7c-.2-.4-.4-.8-.5-1.2l-1.8-.5V8.9l1.8-.5c.1-.4.3-.8.5-1.2l-.8-1.7 1.6-1.6 1.7.8c.4-.2.8-.4 1.2-.5z" />
      <circle cx="10" cy="10" r="2.4" />
    </svg>
  );
}

export function PreviewSettings({
  minimumZoom,
  maximumZoom,
  defaultZoom,
  showModeControl,
  showZoomControls,
  showReviewPanel,
  onMinimumZoomChange,
  onMaximumZoomChange,
  onDefaultZoomChange,
  onShowModeControlChange,
  onShowZoomControlsChange,
  onShowReviewPanelChange,
}: PreviewSettingsProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const firstInputRef = useRef<HTMLInputElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const placePopover = () => {
      const button = buttonRef.current;
      const popover = popoverRef.current;
      if (!button || !popover) return;
      const viewportGap = 12;
      const width = Math.min(328, window.innerWidth - viewportGap * 2);
      const buttonBounds = button.getBoundingClientRect();
      const left = Math.max(
        viewportGap,
        Math.min(
          buttonBounds.right - width,
          window.innerWidth - width - viewportGap,
        ),
      );
      const top = buttonBounds.bottom + 8;
      popover.style.width = `${width}px`;
      popover.style.left = `${left}px`;
      popover.style.top = `${top}px`;
      popover.style.maxHeight = `${Math.max(0, window.innerHeight - top - viewportGap)}px`;
    };
    placePopover();
    window.addEventListener("resize", placePopover);
    window.addEventListener("scroll", placePopover, true);
    return () => {
      window.removeEventListener("resize", placePopover);
      window.removeEventListener("scroll", placePopover, true);
    };
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const focusFrame = requestAnimationFrame(() => firstInputRef.current?.focus());
    const closeFromPointer = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const closeFromKeyboard = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      setOpen(false);
      buttonRef.current?.focus();
    };
    document.addEventListener("pointerdown", closeFromPointer);
    document.addEventListener("keydown", closeFromKeyboard);
    return () => {
      cancelAnimationFrame(focusFrame);
      document.removeEventListener("pointerdown", closeFromPointer);
      document.removeEventListener("keydown", closeFromKeyboard);
    };
  }, [open]);

  return (
    <div className="example-settings" ref={rootRef}>
      <button
        ref={buttonRef}
        className="example-settings-button"
        type="button"
        aria-label="Viewer settings"
        aria-expanded={open}
        aria-controls="preview-settings-popover"
        onClick={() => setOpen((current) => !current)}
      >
        <SettingsIcon />
      </button>
      {open && (
        <div
          ref={popoverRef}
          id="preview-settings-popover"
          className="example-settings-popover"
          role="group"
          aria-label="Preview settings"
        >
          <h2>Preview settings</h2>
          <div className="example-settings-zoom">
            <label>
              <span>Minimum zoom</span>
              <input
                ref={firstInputRef}
                aria-label="Minimum zoom"
                type="number"
                value={minimumZoom}
                placeholder="20"
                onChange={(event) => onMinimumZoomChange(event.currentTarget.value)}
              />
            </label>
            <label>
              <span>Maximum zoom</span>
              <input
                aria-label="Maximum zoom"
                type="number"
                value={maximumZoom}
                placeholder="200"
                onChange={(event) => onMaximumZoomChange(event.currentTarget.value)}
              />
            </label>
            <label>
              <span>Default zoom</span>
              <input
                aria-label="Default zoom"
                type="number"
                value={defaultZoom}
                placeholder="Fit"
                onChange={(event) => onDefaultZoomChange(event.currentTarget.value)}
              />
            </label>
          </div>
          <div className="example-settings-toggles">
            <label>
              <span>Show view controls</span>
              <input
                type="checkbox"
                role="switch"
                checked={showModeControl}
                onChange={(event) => onShowModeControlChange(event.currentTarget.checked)}
              />
            </label>
            <label>
              <span>Show zoom controls</span>
              <input
                type="checkbox"
                role="switch"
                checked={showZoomControls}
                onChange={(event) => onShowZoomControlsChange(event.currentTarget.checked)}
              />
            </label>
            <label>
              <span>Show review panel</span>
              <input
                type="checkbox"
                role="switch"
                checked={showReviewPanel}
                onChange={(event) => onShowReviewPanelChange(event.currentTarget.checked)}
              />
            </label>
          </div>
        </div>
      )}
    </div>
  );
}
