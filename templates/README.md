# templates/

A **ready-to-use starter workbook** — the file users download to begin.

`Example.xlsm` is:

- a workbook with one sheet named **`Sets`** (header in `A1`, e.g. "Search Set Names"; a few example
  names in `A2` downward as a guide),
- with [`../src/modClashLite.bas`](../src/modClashLite.bas) pasted into a standard module (the
  generator), and [`../src/modImportSets.bas`](../src/modImportSets.bas) pasted into a second module
  (optional — imports set names from a Revizto `.vimsst` export),
- saved as a macro-enabled workbook (`.xlsm`).

The `Tests` sheet is created automatically by the macro, so it doesn't need to exist beforehand.

`Testing searchsets.vimsst` is an **example Revizto search-set export** (shareable, no project data).
Use it to try the import: run `ImportSetsFromVimsst` and point the picker at this file — it fills the
`Sets` sheet so you can see the round-trip without needing your own export first.

The starter must contain **no real project data** — only placeholder example names. The live working
workbook is git-ignored and must never be committed. Because `*.xlsm` is git-ignored, force-add the
starter when updating it:

```
git add -f templates/Example.xlsm
```
