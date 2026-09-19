# Source and modifications

DocJS is a maintained derivative of [docxjs](https://github.com/VolodymyrBaydalka/docxjs),
whose npm package is `docx-preview`. This repository retains its Git history and
Apache-2.0 license. The starting point for the React extraction is version 0.4.0,
commit `191d3e0db009da578fbe4da70d55305cd8d50226`.

The React preview, review interaction, and measured pagination were developed in
Falcon's `preview-poc`. The extraction uses checkpoint
`6892a1c0e619685660aa66cd48a6843a865b95eb`, including its styling and rendering
corrections. Those files are project additions, separate from the inherited
upstream engine. The component uses the engine source in this repository.

## Rendering corrections

The extraction initially keeps the POC adapters so that moving the code does not
also require rewriting layout. They preserve document-derived fonts, dimensions,
spacing, numbering, section geometry, page fields, and measured pagination.

The last POC correction set includes:

- Minimum line spacing and property-level inheritance without omitted values
  erasing inherited values.
- Collapsing eligible empty continuous-section carrier paragraphs while keeping
  their model identities and section metadata.
- Excluding footer free space from occupied-height calculations during incremental
  pagination.
- Neutralizing cached Word page-break markers while preserving authored breaks.

`ignoreLastRenderedPageBreak` remains false in the viewer's engine options. The
geometry adapter neutralizes cached markers itself because the inherited option
also affects continuous-section grouping. Changing that option requires checking
both grouping and pagination.

## Maintenance boundary

The engine remains independent of React. Consumers use the React entry point for
the viewer and provide document bytes and review items. Rendering adapters may
move into the parser or layout engine when equivalent behavior has been verified.

Useful upstream fixes can be ported selectively. This derivative does not promise
compatibility with upstream internal structures or automatic wholesale merges.
Keep attribution and the original license when distributing it. This extraction
does not establish a new package name or publish a release.
