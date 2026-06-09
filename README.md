# Revizto Clash-Test Generator

Turn a **list of search-set names** into a full matrix of Revizto **clash tests** (`.vimctst`) —
one button, instead of creating and naming hundreds of tests by hand.

A coordination project needs a clash test for every pair of disciplines/elements. Setting those up
in Revizto is slow and repetitive. This tool takes your list of search sets, builds every pairing,
and writes a `.vimctst` you import straight into Revizto.

> ⚠️ **Unofficial.** The `.vimctst` format is **reverse-engineered**, not documented or supported by
> Revizto, and this project is **not affiliated with or endorsed by Revizto**. The format can change
> between Revizto versions. **Always keep backups and work on copies.** Use at your own risk.

---

## What it does

```
Sets sheet (your list of set names)
        │  GenerateClashTests
        ▼
Tests sheet (every pairing, 1v1 / 2v1 / 2v2 / …)  +  Clash Tests.vimctst
        ▼
import into Revizto
```

- **In:** one column of search-set names (the `Sets` sheet).
- **Out:** a `.vimctst` containing one clash test for **every lower-triangular pair, including
  self-pairs** (so `N` sets → `N·(N+1)/2` tests), plus a `Tests` sheet showing the matrix.
- **No codes, priorities, or clearances** — just names in, tests out. Set those up in Revizto after.

## Why it works

On import, **Revizto re-matches each test's two sides to your project's search sets by the set
*name*** — the stored GUID and folder path are cosmetic. So the tool only needs your set *names*; it
fabricates the GUIDs. A name that doesn't match a set imports with a warning and a "re-select from
list" prompt — so typos are obvious and harmless, never destructive.

## Use it

1. **Set up the workbook** (once): a sheet named **`Sets`** with a header in `A1` and your search-set
   names in `A2` downward. Then `Alt+F11` ▸ `Insert ▸ Module`, paste
   [`src/modClashLite.bas`](src/modClashLite.bas). To pull names straight from Revizto instead of
   typing them, also paste [`src/modImportSets.bas`](src/modImportSets.bas) (see below).
2. **Run** `GenerateClashTests` (`Alt+F8`, or F5 from the editor).
3. It writes the **`Tests`** sheet (preview of every pair) and saves **`Clash Tests.vimctst`** next to
   the workbook.
4. **Import** the `.vimctst` into Revizto. The set names must match your real Revizto sets.

### Don't want to type the names? Import them from Revizto

Paste [`src/modImportSets.bas`](src/modImportSets.bas) into a second module and run
`ImportSetsFromVimsst`. It shows a file picker — point it at a **`.vimsst`** search-set export from
Revizto, and it fills the **`Sets`** sheet for you:

- **column A** — the set name (exactly as stored; this is what the generator uses),
- **column B** — the Revizto folder each set lived under (reference only — the generator ignores it;
  handy for spotting and deleting sets you don't want).

Existing names in column A are replaced (you're asked to confirm first). Then carry on at step 2.
Because names must match your real Revizto sets exactly, importing them is also the safest way to
avoid typos.

To try it without your own export, point the picker at the bundled example
[`templates/Testing searchsets.vimsst`](templates/Testing%20searchsets.vimsst).

> **OneDrive gotcha:** the macro saves next to the workbook via `ThisWorkbook.Path`. For a workbook in
> OneDrive/SharePoint that path can be an `https://…` URL that file I/O can't use (run-time error 52).
> Run from a **local, non-synced folder** (e.g. `C:\Temp\`) and copy the result back.

## Repository layout

```
src/modClashLite.bas              the generator — paste into the workbook
src/modImportSets.bas             optional: import set names from a Revizto .vimsst export
docs/vimctst-format.md            reverse-engineered .vimctst format notes
templates/Example.xlsm            ready-to-use starter workbook (both macros pasted in)
templates/Testing searchsets.vimsst  example .vimsst export to try the import with
```

## Status

Private / work in progress. The generator works. **Before any public release:** confirm Revizto are
comfortable with it, clear any employer/IP questions, and add a clean starter workbook with no
project data.

## License

[MIT](LICENSE) — pending IP review (see Status).
