# Selected fork integration — 2026-09-21

Adapted useful source changes from the [fork network review](fork-network-review-2026-09-21.md) onto `667e240`, preserving the selective upstream fixes in `b63a493`. These are local adaptations, not whole-branch merges or copies of published bundles.

## Integrated sources

| Source | Integrated behavior and adaptations |
| --- | --- |
| [loadfix `20eca4e290`](https://github.com/loadfix/docxjs/commit/20eca4e290f43948b7a3f3d3e006d562806786eb) | Contain HTML altChunks with an empty iframe sandbox, preserving the existing option and HTML display rather than deleting the feature. Escape numbering label literals while preserving our per-level counter formatting. Use null-prototype lookup maps and reject prototype-mutating keys/inherited properties in deep merges. Retain our existing hyperlink allowlist. Do not import restrictive color sanitization. |
| [loadfix `0e59e7c19b`](https://github.com/loadfix/docxjs/commit/0e59e7c19bb18354b615e9582e7396376de3b49a) | Render vertically aligned run children once, preventing repeated footnote reference collection and other stateful rendering. The continuous-numbering portion is not imported. |
| [loadfix `1e4352be64`](https://github.com/loadfix/docxjs/commit/1e4352be6406a5857d8caa87c833a893feaade3c) | Remove Word's numeric index suffix from VML fill/stroke colors, retaining named colors and `none`. |
| [klarso `f9387a6bddc5`](https://github.com/klarso-gmbh/docxjs/commit/f9387a6bddc5) | An empty `autoHyphenation` element means enabled. Explicit false and absent settings retain their distinct behavior. |
| [klarso `369f4ed4c0cb`](https://github.com/klarso-gmbh/docxjs/commit/369f4ed4c0cb) | Identify document defaults explicitly, then omit ordinary styles with no ID from stylesheet emission. |
| [klarso `2b55273aa48e`](https://github.com/klarso-gmbh/docxjs/commit/2b55273aa48e), [klarso `d08207d116ef`](https://github.com/klarso-gmbh/docxjs/commit/d08207d116ef) | Re-scan cloned paragraph continuations to handle multiple page breaks. Preserve parsed break nodes for repeat rendering. Keep our direct false precedence, breakPages switch, list marker suppression and hidden-content checks. Section properties remain on the final fragment. No measured-pagination dependency is introduced. |
| [M-Busk `377c3d5de4`](https://github.com/M-Busk/docxjs/commit/377c3d5de4a384237c1d6894bfc68c337e95628a) | Parse comment range start/end markers inside hyperlinks. |
| [flowkscai `8ceaf69354`](https://github.com/flowkscai/docxjs/commit/8ceaf69354eda60a172c441fe7570ac2c35b6ed9) | Resolve same-property-set highlighting ahead of shading regardless of XML order. Explicit `highlight=none` preserves shading from that property set. |
| [Luo-Studio `c2f962a1fd`](https://github.com/Luo-Studio/docxjs/commit/c2f962a1fd) | Preserve and render cached simple-field children, including formatting and nested fields. Read the actual `fldLock` attribute. No speculative PAGE/NUMPAGES/SECTIONPAGES evaluation. |
| [fcapurso `8f790d845499`](https://github.com/fcapurso/docxjs/commit/8f790d845499) | Resolve omitted paragraph styles through the declared default paragraph style for classes, tabs, numbering and pageBreakBefore. Do not assume a style named `Normal` is necessarily the default. |
| [DmitrySharabin `3d7c2932659f`](https://github.com/DmitrySharabin/docxjs/commit/3d7c2932659f) | `renderAsync` waits for font readiness and measures experimental tabs after attachment. Retain sequential measurements rather than importing the previously rejected batch algorithm. Detached `renderDocument` keeps its legacy delayed pass; disconnected tab spans are skipped. |
| [Questra `034fbcb23a`](https://github.com/Questra-ai/docxjs/commit/034fbcb23a313d6f0be78ca34df911dbde77c9c8) | Infer header/last-row and column/band classes when conditional classes are absent. Respect explicit all-false conditional formatting, grid offsets, spans, direct/style band sizes and table look flags. Apply conditional fills to cells with direct-child selectors so nested table cells are not selected. Enable the corresponding table look modifiers; explicit boolean attributes override packed flags. |

Source authors remain credited through the pinned commit links. All distributions are rebuilt from the combined source.

## Validation

- TypeScript and production Rollup build pass.
- All **102 Chromium browser tests pass** against both production UMD and minified UMD bundles.
- Added **30 browser regressions**. Against the original `667e240` bundle, 17 fail and 13 pass; passing cases check behavior that must remain intact.
- The altChunk test demonstrates script execution against the baseline and its prevention by the sandbox. It also verifies that HTML remains present and the opt-out still omits it.
- Tests cover actual immediate tab geometry, three-page splitting, consecutive breaks, numbered continuations, repeat rendering, cached field formatting, table fills, nested cells, skipped columns, explicit flags and direct cell formatting.
- Additional utility checks exercise null-prototype maps, malicious merge keys and CSS string escaping.
- Existing HTML snapshots were reviewed and updated for default paragraph classes, removal of empty page-break spans and conditional table selector changes. Snapshot changes are not a claim of pixel-perfect Word fidelity.
- The existing Rollup external-JSZip warning remains.

## Deliberately excluded

The following changes still require a separate implementation/evaluation; the review did not establish them as safe imports:

- Whole measured-pagination engines, table-row/line splitting, keep-next layout, repeated headers, virtual scrolling and worker/snapshot APIs.
- Live page-field evaluation without section numbering/restart/format semantics.
- Forcing continuous footnote numbering without honoring restart policies and note-list labels together.
- SVG viewBox clamping without stroke/fractional-boundary visual coverage, linked textbox flow and broad drawing placement rewrites.
- Full highlight-versus-shading inheritance across separate style layers; this integration fixes precedence within a parsed property set.
- The four fully deferred upstream PRs and the unsafe portions of partially integrated PRs recorded in the prior review.

Klarso's pinned branch remains a useful source for the next layout phase. The current integration takes its bounded correctness fixes while keeping our API and existing regression coverage.
