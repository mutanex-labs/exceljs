# Upstream characterization

This suite records observable behavior at upstream baseline `5bed18b45e824f409b08456b59b87430ded023ab`. It does not modify runtime code, dependencies, or package metadata. These tests are kept outside `spec/integration` so expected upstream defects do not make the ordinary integration suite permanently red.

## Categories

- `upstream-pass`: the baseline is expected to satisfy the assertion.
- `known-upstream-bug`: a separate probe records a currently observed incorrect result or error. These probes do not join the passing Mocha suite.
- `baseline-limitation`: the current writer/reader loses or rewrites unsupported or partially supported package content, or exposes a documented default behavior.
- `stress/invariant probe`: a workload or transport perturbation that records correctness and failure surface without timing or RSS assertions.

The category appears in each Mocha test title or in the machine-readable probe entry.

## Run

Run the pass and mutation characterization tests:

```text
node node_modules/mocha/bin/_mocha --require spec/config/setup spec/characterization/baseline-pass.spec.js spec/characterization/mutation-preservation.spec.js
```

Run the known-bug, transport, and repeated-race observations:

```text
node spec/characterization/run-streaming-probes.js --runs=10 --race-runs=20
```

The three frozen regression inputs live under `spec/characterization/fixtures/`; the runner checks their hashes and verifies that the date/shared-string and race ZIPs contain identical OOXML part content. Repeated probe outcomes and runtime metadata are stored in `.temp/characterization/streaming-adversarial.json`. Mutation tests store normalized package differences under `.temp/characterization/mutation-preservation.json` and their generated XLSX inputs/outputs in the same ignored directory.

No timing or RSS assertion is part of these tests. The 100k-row writer case verifies row count and values only. The existing import measurements remain unchanged baseline data; subsequent reports call this metric `fresh-process require/import` because page and filesystem cache state is uncontrolled.

## Coverage and fixture provenance

| Coverage | Category | Fixture/source |
|---|---|---|
| XLSX read/write, multiple worksheets, styles, merge, formula and cached result, rich text, named range, validation, hyperlink, shared strings, hidden sheets, conditional formatting, PNG image | `upstream-pass` | Generated in test code; image is the existing upstream `spec/integration/data/image.png` fixture |
| 1900 and 1904 date systems | `upstream-pass` | Generated workbooks; 1904 behavior also has upstream issue 219 coverage |
| WorkbookReader and WorkbookWriter round-trip | `upstream-pass` | Small generated workbooks |
| CSV read/write default behavior | `baseline-limitation` | Generated CSV; the date field is emitted as ISO text and read back as a string |
| Large WorkbookReader input | `stress/invariant probe` | Existing upstream `spec/integration/data/huge.xlsx`; 100k data rows plus header |
| 100k-row streaming write | `stress/invariant probe` | Generated in memory; output is read back and row values are counted; no performance measurement |
| Streaming date/styles/sharedStrings issue 1328 | `known-upstream-bug` | Existing 4.7 KB upstream `spec/integration/data/dateIssue.xlsx` plus frozen 2.0 KB seven-part OOXML fixture `fixtures/minimal-styles-date-sharedstrings.xlsx`: one shared-string header and one styled date |
| Rich-text sharedStrings issue PR 1431 | `known-upstream-bug` | Frozen one-row baseline WorkbookWriter fixture `fixtures/minimal-richtext-sharedstrings.xlsx`: one rich-text value and one plain shared string |
| `WorkbookReader` worksheet metadata race | `known-upstream-bug` | Frozen `fixtures/minimal-workbookreader-order-race.xlsx`; its part contents exactly match the date/shared-string fixture and only the valid ZIP entry order differs. The exact upstream `WorkbookReader Serialise xlsx file` workload also repeats in fresh processes |
| Chunk boundaries, entry order, slow consumer/backpressure, abort, truncation, deterministic repeats | `stress/invariant probe` | One generated logical workbook; transport and ZIP entry order are varied without changing cell content |
| Single-cell mutation and OOXML part diff | `baseline-limitation` | Generated synthetic workbook and existing upstream chart-sheet/comment fixtures |
| VBA part, external links, custom XML, vendor XML, theme, relationships | `baseline-limitation` | Programmatically injected OOXML parts; the VBA payload is a placeholder, not an executable macro project |
| Chart sheets, drawings, comments/VML, printer settings, third-party package parts | `baseline-limitation` | Existing upstream fixtures or synthetic relationship-linked parts |

The synthetic unsupported-part package is an OOXML preservation probe, not a claim that its placeholder VBA payload is a valid XLSM macro. Existing upstream fixtures are referenced in place and are not copied into new tracked fixtures.

The mutation report distinguishes dropped parts from re-emitted parts with new names. In the issue 880 fixture, the original comment/VML part names disappear while new comment/VML parts appear; printer settings are dropped. That inventory alone does not claim that the comment text itself was lost.

## Baseline observations

The pass/mutation command currently reports **13 passing tests**. The probe results below are from the same upstream SHA and are stored in the JSON files described above.

- Issue 1328 reproduces in 10/10 reads on the upstream fixture. The minimal generated fixture also reproduces in 10/10 reads: the header is `{sharedString: 0}`, while the styled date cell decodes to `2020-11-20T00:00:00.000Z`.
- In the final 10-run probe, PR 1431's frozen WorkbookWriter archive returns shared-string indices in 10/10 full-buffer reads. With identical OOXML parts reordered so workbook metadata, relationships, and shared strings precede the worksheet, it decodes rich text/plain text in 10/10. On the original ZIP, chunks 1 and 7 decode correctly in 10/10, chunk 64 is mixed (9/1), chunk 1024 emits indices in 10/10, and the irregular pattern decodes in 10/10. A separate invocation previously had one successful full-buffer read, so this symptom is itself timing-sensitive.
- For the same 21-row WorkbookWriter archive, chunk sizes 1, 7, and 64 yield correct rows in 10/10 runs; chunk size 1024 yields shared-string indices in 10/10; the repeating 1/7/64/3 pattern is mixed (9/1).
- For the same Workbook.xlsx `writeBuffer` contents, chunk size 1 passes 10/10, chunk size 64 has one `sheets` TypeError in 10, and chunk size 1024 has nine `sheets` TypeErrors in 10.
- Original and reverse ZIP entry orders pass 10/10. The frozen valid-order permutation keeps every OOXML part byte-identical but places `workbook.xml` after the worksheet; it raises the `sheets` TypeError in 20/20 repeated runs. The exact upstream race test passes 20/20 in isolated fresh processes; Phase 1 full integration repeats failed it in 3/5 runs, so the flaky symptom depends on suite context/order.
- Slow async-iterator consumption produces correct values in 4/5 runs and shared-string indices in 1/5. Destroying input when row 5 is observed results in 13 buffered rows being emitted, then no terminal event within 1.5 seconds in 5/5 runs. Truncated input reports `FILE_ENDED` in 10/10 runs.
- Mutation keeps the formula cached result `42`; the generated fixture's image/drawing, custom-marked theme, and styles remain byte-identical. The synthetic VBA placeholder, external-link parts, custom XML, and vendor XML parts are dropped. The chart-sheet fixture loses its chart, chart sheet, drawing, and corresponding relationships. The issue 880 fixture drops printer settings and re-emits comment/VML parts under new names while worksheet parts and relationships are rewritten.

The synthetic VBA payload is not a valid macro project, and the upstream fixture set contains no `.xlsm` asset. The probe therefore confirms that the writer drops a `vbaProject.bin` package part, but it does not characterize preservation of an executable macro project.
