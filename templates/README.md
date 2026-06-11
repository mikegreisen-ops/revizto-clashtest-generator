# templates/

A **ready-to-use starter workbook** — the file users download to begin.

`Example.xlsm` is:

- an **`Instructions`** sheet (the workbook opens here) with the steps and a **button for each macro**,
- a **`Sets`** sheet (just the header in `A1`, e.g. "Search Set Names"; **no example names** — filled via
  `ImportSetsFromVimsst` or by typing/pasting),
- a **`Tests`** sheet that **ships ready to use** (formatted headers incl. the `Type` column, no rows) so
  someone with an existing test list can paste it straight in,
- with [`../src/modClashLite.bas`](../src/modClashLite.bas) and
  [`../src/modImportSets.bas`](../src/modImportSets.bas) pasted into standard modules,
- saved as a macro-enabled workbook (`.xlsm`).

`BuildTestList` and `ImportClashTests` also (re)create the `Tests` sheet on demand.

`Testing searchsets.vimsst` is an **example Revizto search-set export** (shareable, no project data).
Use it to try the name import: run `ImportSetsFromVimsst` and point the picker at this file — it fills
the `Sets` sheet so you can see the round-trip without needing your own export first.

## Rebuilding the starter

The `.bas` files in `../src/` are the source of truth; this workbook holds its own embedded copy.

- After changing a `.bas`, re-sync just the modules with
  [`../tools/sync-xlsm.ps1`](../tools/sync-xlsm.ps1) (avoids leaving stale duplicate modules behind).
- To rebuild the **whole** workbook — modules **and** the Instructions / Sets / Tests sheets and buttons
  — run [`../tools/setup-starter.ps1`](../tools/setup-starter.ps1).

Both need Excel + the one-time *Trust access to the VBA project object model* setting, with the workbook
closed. Close stray Excel instances first — a lingering one can clobber the save.

## Keep it clean

The starter must contain **no real project data** — the `Sets` sheet ships empty (header only) and the
only sample data is the shareable example `.vimsst`. The live working workbook is git-ignored and must
never be committed. Because `*.xlsm` is git-ignored, force-add the starter when updating it:

```
git add -f templates/Example.xlsm
```
