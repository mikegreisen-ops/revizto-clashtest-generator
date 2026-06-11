# templates/

A **ready-to-use starter workbook** — the file users download to begin.

`Example.xlsm` is:

- a workbook with one sheet named **`Sets`** (just the header in `A1`, e.g. "Search Set Names"; **no
  example names** — you fill column A with `ImportSetsFromVimsst` from the bundled example `.vimsst`, or
  by typing/pasting your own),
- with [`../src/modClashLite.bas`](../src/modClashLite.bas) pasted into a standard module (the
  `BuildTestList` / `ImportClashTests` / `ExportClashTests` macros), and
  [`../src/modImportSets.bas`](../src/modImportSets.bas) in a second module (optional — imports set
  names from a Revizto `.vimsst` export),
- saved as a macro-enabled workbook (`.xlsm`).

The `Tests` sheet (the editable list of pairs with Tol / Clr / Priority / Stamp columns) is created
automatically by `BuildTestList` or `ImportClashTests`, so it doesn't need to exist beforehand.

`Testing searchsets.vimsst` is an **example Revizto search-set export** (shareable, no project data).
Use it to try the name import: run `ImportSetsFromVimsst` and point the picker at this file — it fills
the `Sets` sheet so you can see the round-trip without needing your own export first.

## Keeping the embedded macros current

The `.bas` files in `../src/` are the source of truth; this workbook holds its own embedded copy.
After changing a `.bas`, re-sync the starter with [`../tools/sync-xlsm.ps1`](../tools/sync-xlsm.ps1)
(it re-imports the modules into `Example.xlsm`) rather than pasting by hand — that avoids leaving stale
duplicate modules behind.

## Keep it clean

The starter must contain **no real project data** — the `Sets` sheet ships empty (header only) and the
only sample data is the shareable example `.vimsst`. The live working workbook is git-ignored and must
never be committed. Because `*.xlsm` is git-ignored, force-add the starter when updating it:

```
git add -f templates/Example.xlsm
```
