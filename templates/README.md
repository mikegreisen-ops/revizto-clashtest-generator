# templates/

This folder will hold a **ready-to-use starter workbook** — the file users download to begin.

It only needs to be:

- a workbook with one sheet named **`Sets`** (header in `A1`, e.g. "Search Set Names"; a few example
  names in `A2`–`A4` as a guide),
- with [`../src/modClashLite.bas`](../src/modClashLite.bas) pasted into a standard module,
- saved as a macro-enabled workbook (`.xlsm`), e.g. `clash-generator-starter.xlsm`.

The `Tests` sheet is created automatically by the macro, so it doesn't need to exist beforehand.

Make sure the starter contains **no real project data** — only placeholder example names. The live
working workbook is git-ignored and must never be committed. When the starter is ready, force-add it:

```
git add -f templates/clash-generator-starter.xlsm
```
