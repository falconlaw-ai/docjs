# DocJS

React document preview component for PDF, DOCX, and Markdown originals, with a working-DOCX review view. The repository
also owns the underlying DOCX rendering engine and retains the `docx-preview` package name.

## Legit DX

- Issue tracker: Linear
  - Linear team: DocJS
  - Issue prefix: DOC
- Forge: GitHub
  - Default branch: master
- QA contract: docs/quality-assurance.md
- QA delegation: one fresh general-purpose agent

### Domains

- React preview
  - Paths: `src/react/**`, `src/review/**`, `example/**`, and `tests/preview/**`
  - Documentation: `README.md` when the public component API, integration contract, example, or supported behavior
    changes
  - Verify: `pnpm typecheck`, `pnpm build-prod`, `pnpm test:package`, `pnpm test:review`, `pnpm example:build`, then
    `pnpm test:preview`
  - Restrictions: Keep file conversion, transport, caching, persistence, credentials, and export in the consumer app.
- DOCX engine
  - Paths: engine code under `src/**` outside `src/react/**` and `src/review/**`, plus `tests/render-test/**` and
    `tests/extended-props-test/**`
  - Documentation: `README.md` when the root entry point or documented rendering behavior changes
  - Verify: `pnpm typecheck`, `pnpm build-prod`, then
    `pnpm exec karma start karma.conf.cjs --single-run --browsers ChromeHeadless`
  - Restrictions: Preserve `renderAsync` as the stable engine entry point; treat parser and renderer internals as
    implementation details.
- Tooling and packaging
  - Paths: `package.json`, `pnpm-lock.yaml`, `pnpm-workspace.yaml`, Rollup, TypeScript, Karma, Playwright, and
    `.github/**`
  - Verify: run `pnpm typecheck`, `pnpm build-prod`, and each affected test or example-build script
  - Restrictions: Keep generated `dist/**`, `example/dist/**`, Playwright reports, and copied example assets out of pull
    requests.

## Source of truth

Read [README.md](README.md) before changing the public API, rendering behavior, review model, example integration,
security boundaries, or verification commands. Treat package scripts and the Rollup, TypeScript, Karma, Playwright,
and Vite configuration as the source of truth for build and test behavior.

## Workflow

Use the Legit `workflow` skill when work starts, resumes, redirects, pauses, or asks what comes next. A specifically
named workflow stage selects that stage directly.

## Hard rules

- Keep the reusable React component independent of any consumer backend. PDF or Markdown conversion to the working DOCX
  happens outside this package; the example may simulate that boundary with its deterministic fake backend.
- Do not add SuperDoc, Yjs, Hocuspocus, conversion credentials, or application-specific transport and persistence code.
- Preserve the original document in Original mode. Review and Final render only from the supplied working DOCX and must
  not replace or mutate the source file.
- Keep document rendering browser-only and preserve the resource-isolation and cleanup guarantees documented in the
  README.
- Add or update representative fixtures and focused tests when rendering, pagination, preparation, navigation, review
  projection, or package behavior changes.
- Use repository fixtures or synthetic documents in tests. Do not add customer documents or confidential content.
- Keep generated build output and copied assets out of pull requests.
