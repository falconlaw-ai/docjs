# DocJS document preview

A React viewer for original PDF, DOCX, and Markdown documents, with a working-DOCX
review view and externally supplied comments and redlines. This repository owns
the DOCX engine source and retains the `docx-preview` package name. The React
extraction is not a published npm release.

The viewer ports Falcon's preview POC, including its styling, typography and
geometry corrections, measured pagination, and review navigation. See
[UPSTREAM.md](UPSTREAM.md) for source checkpoints and maintenance boundaries.

## Run the example

Use Node 24 and pnpm:

```sh
pnpm install
pnpm build-prod
pnpm example:dev
```

The example uses prepared fixtures and a deterministic fake backend. It shows
Original, Review, and Final, preparation failure and retry, review item updates,
and custom cards. It performs no PDF or Markdown conversion. Uploaded documents
can be previewed locally; a fixture mapping is required for the example's prepared
review flow.

## React integration

Import the component and its stylesheet from separate entry points:

```tsx
import { useState } from "react";
import {
  DocumentPreview,
  type PreviewDocument,
  type PreviewMode,
  type ReviewItem,
} from "docx-preview/react";
import "docx-preview/react/styles.css";

export function ReviewDocument({
  original,
  items,
}: {
  original: PreviewDocument;
  items: readonly ReviewItem[];
}) {
  const [mode, setMode] = useState<PreviewMode>("original");
  return (
    <DocumentPreview
      original={original}
      mode={mode}
      onModeChange={setMode}
      items={items}
    />
  );
}
```

`original` contains a stable `id`, a `revision`, a `format` of `docx`, `pdf`, or
`md`, and a browser `File`. Change the revision when the document changes. Keep
file identity stable across review-item updates. The viewer is browser-only and
uses React 19. Its parent must give it usable width and height.

A source DOCX can serve as its own working document. For PDF and Markdown, supply
`preparation` and `onRequestPreparation`:

- `unavailable` means a working document has not been supplied.
- `pending` contains the original's `{ id, revision }` as `identity`.
- `ready` contains that identity and the prepared DOCX `file`.
- `error` contains that identity and a displayable `message`.

The callback receives the original document and asks the application to prepare
or retry it. The application owns the request, cancellation, caching, and state.
The viewer ignores preparation for another identity. Original stays visible
while preparation or the first working-document render is pending. Failure leaves
the original usable. See [the example](example/src/App.tsx) and
[its fake backend](example/src/fakeBackend.ts) for a complete integration.

## Review items and Final

Supply an immutable item array. Each item has a stable external `id`, a `kind`
of `comment` or `redline`, and an `anchor` with `quote` plus optional `prefix`,
`suffix`, and `paragraphId`. Comments use `body`; redlines use `replacement`.
Optional `metadata` carries application data for custom cards.

Anchors resolve against the working DOCX's original text. Paragraph IDs are
scoped to that document revision. The resolver understands its ordered `pN`
identifiers and native Word paragraph IDs where available, and checks the quoted
text rather than trusting an ID alone. Repeated quotes need enough context to
resolve uniquely. Unsupported ranges, stale anchors, and overlapping redlines
produce visible diagnostics. Cross-paragraph ranges and ranges crossing tabs or
line breaks are not supported in this extraction.

Review displays comments and proposed replacements. **Final hides comments and
shows every supplied resolvable, nonconflicting redline as replacement text.** It
does not approve items, change the source bytes, or export a document. Consumers
that need an approved subset must filter the supplied array. Final also retains
the POC's accepted-looking projection of native Word tracked changes.

Item updates reuse the document source and patch affected paragraphs, then
repaginate the affected section tail. They do not rebuild a DOCX ZIP. Switching
modes may reparse and repaginate the prepared document; it does not reconvert it.
A failed update preserves the last committed rendering and reports a refresh
error through `onRefreshError`.

## Custom cards and appearance

`renderReviewItem` receives `{ item, entity, selected, diagnostic }`. Return the
card body and any application-owned action controls. The viewer retains the outer
selection and navigation target, including repeated clicks, keyboard activation,
and document-to-card scrolling. Use application callbacks from custom controls
for business actions such as approval; the viewer has no approval policy.

The default shell preserves the POC's grey background, white pages, small page
corners, zoom controls, and review-card borders. Use `className` and `style` to
adapt the shell. DOCX styles live in Shadow DOM so application CSS does not replace
source typography. Zoom scales rendered pages rather than changing logical
pagination.

## PDF resources and document isolation

Pass `pdfAssets` with `workerUrl` and `resourceBaseUrl` pointing to locally hosted
resources from the same `pdfjs-dist` version used by the package. The resource
base contains `cmaps/`, `standard_fonts/`, `wasm/`, and `iccs/`. The worker is
`build/pdf.worker.min.mjs` from that installation. The example's asset-copy script
demonstrates deployment without a CDN or Vite-specific imports in the library.

The renderers restrict document resources, skip Markdown raw HTML, disable DOCX
HTML chunks, and isolate DOCX styles. Hosts must also enforce a Content Security
Policy that limits images, fonts, workers, and other embedded resources to the
intended local, blob, or data sources. The example includes that policy. Do not
assume post-render cleanup alone prevents every resource fetch in an arbitrary
host. External links remain explicit user actions.

The library owns renderer cleanup and superseded rendering work. Applications own
document retrieval, transport subscriptions, persistence, conversion credentials,
and export. The package includes no SuperDoc, Yjs, Hocuspocus, or backend worker.

## Rendering limits and verification

The extraction preserves the POC's corrections. It does not claim exact Word
layout. Font availability affects wrapping, and complex fields, multi-column
flow, floating objects, widow/orphan rules, and notes still have limitations.

The fresh reference POC rendered the consulting fixture as five Letter pages and
the complex MSA as 17 pages; its source PDF has 18 pages. Page counts are comparison
evidence, not a reason to override document fonts, spacing, margins, or authored
breaks. Fixture provenance is recorded in
[the fixture notes](tests/preview/fixtures/README.md).

```sh
pnpm build-prod
pnpm typecheck
pnpm test:package
pnpm test:review
pnpm exec karma start karma.conf.cjs --single-run --browsers ChromeHeadless
pnpm example:build
pnpm test:preview
```

The inherited Karma/Jasmine suite exercises the engine. Review-model tests cover
anchor projection and incremental changes. Browser tests mount the built package
and exercise rendering, preparation, mode changes, and navigation. Browser tests
require Chrome and use the local example server.

## Engine entry point

The root entry remains available for non-React use:

```ts
import { renderAsync } from "docx-preview";
await renderAsync(file, container, undefined, { renderChanges: true });
```

The React viewer calls this repository's parser and renderer internally. Engine
internals are not the React integration contract. Consult the source types for
engine options, and preserve [LICENSE](LICENSE) and upstream attribution when
distributing the derivative.
