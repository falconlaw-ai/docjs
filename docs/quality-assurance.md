# Quality assurance

This policy defines post-review QA for DocJS. Legit DX owns workflow state, retries, repair loops, and publication
sequencing.

## Applicability and coverage

QA is required when a change can alter document parsing or rendering, generated HTML or styles, pagination, preparation,
review projection, browser behavior, the public API, package output, or the example integration. Documentation-only
changes and behavior-preserving test or tooling changes are inapplicable when they cannot affect a runtime path.

Use the smallest representative fixture for the changed behavior, then run the relevant complete suite. Rendering
changes need a repository fixture or synthetic reproduction and a focused assertion. A missing fixture may be bypassed
only with explicit user approval that records the uncovered behavior, alternative checks, and residual risk.

## Delegation and execution

Use one fresh general-purpose agent. Run QA from the reviewed worktree against the exact reviewed commit.

The local worktree is the only approved environment. Use the package scripts and checked-in browser configuration for
startup and execution. QA must not depend on production systems, external conversion services, or customer documents.
Chrome is required by both the Karma and Playwright suites.

## Checks and evidence

Run `pnpm typecheck` and `pnpm build-prod` first. For DOCX engine changes, run
`pnpm exec karma start karma.conf.cjs --single-run --browsers ChromeHeadless`. For React preview or review-model changes,
run `pnpm test:package`, `pnpm test:review`, `pnpm example:build`, and `pnpm test:preview`. Add a focused manual browser
check when the automated assertions cannot establish the visual behavior.

Store evidence under `.legit-dx/runs/<run>/evidence/attempt-<n>/<check-id>/`. Record each command, exit status, and
relevant output. Store screenshots or traces for browser failures and manual visual checks. Evidence must identify the
reviewed commit and exclude confidential or customer-provided content.

Use repository fixtures or synthetic reproductions only. QA creates no durable external data; close transient browser
processes when the run finishes.

## Limits and outcomes

Allow at most two complete QA attempts, one repair batch, and one infrastructure retry within an attempt. Stop after 20
minutes or earlier when Chrome is unavailable, the environment is unsafe, or a required reproduction cannot be stored
safely.
