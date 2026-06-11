# Revizto Clash-Test Generator

Turn a **list of search-set names** into a full matrix of Revizto **clash tests** (`.vimctst`), with
per-test tolerance / clearance / priority — instead of creating, naming and configuring hundreds of
tests by hand. It also reads an existing `.vimctst` **back** into Excel, so you can bulk-edit settings
and re-export.

A coordination project needs a clash test for every pair of disciplines/elements. Setting those up in
Revizto is slow and repetitive. This tool builds every pairing into an editable `Tests` sheet, lets
you drive the settings from your own matrix, and writes a `.vimctst` you import straight into Revizto.

> ⚠️ **Unofficial.** The `.vimctst` format is **reverse-engineered**, not documented or supported by
> Revizto, and this project is **not affiliated with or endorsed by Revizto**. The format can change
> between Revizto versions. **Always keep backups and work on copies.** Use at your own risk.

---

## What it does

```
 Sets sheet                 Tests sheet                       Clash Tests.vimctst
 (set names)  ──BuildTestList──▶  (one row per pair,    ──ExportClashTests──▶  (import into Revizto)
                                   editable Tol/Clr/Prio)
                                        ▲
 a .vimctst  ──ImportClashTests────────┘   (read existing tests back in to edit)
```

Three macros, one editable `Tests` sheet in the middle:

| Macro | Direction |
|---|---|
| **`BuildTestList`** | `Sets` → `Tests` sheet — every lower-triangular pair **including self-pairs** (`N` sets → `N·(N+1)/2` tests) |
| **`ImportClashTests`** | `.vimctst` → `Tests` sheet — decode an existing export to edit it |
| **`ExportClashTests`** | `Tests` sheet → `Clash Tests.vimctst` — write the file to import into Revizto |

So you can work **from scratch** (BuildTestList → adjust → ExportClashTests) or **round-trip** an
existing set (ImportClashTests → adjust → ExportClashTests).

## Why it works

On import, **Revizto re-matches each test's two sides to your project's search sets by the set
*name*** — the stored GUID and folder path are cosmetic. So the tool only needs your set *names*; it
fabricates the GUIDs. A name that doesn't match a set imports with a warning and a "re-select from
list" prompt — so typos are obvious and harmless, never destructive. (Priority and stamp bind the same
way — by value/code, not GUID.)

## The `Tests` sheet

`BuildTestList` and `ImportClashTests` both write the same sheet; `ExportClashTests` reads it.

| Col | Header | Meaning | Default |
|---|---|---|---|
| A | **Test Name** | `"<Set A> vs <Set B>"` — rename freely, it becomes the test's name | — |
| B | **Tol (mm)** | tolerance distance (mm) | `25` |
| C | **Clr (mm)** | clearance distance (mm) | *(blank)* |
| D | **Priority** | `Trivial` / `Minor` / `Major` / `Critical` / `Blocker` | `Minor` |
| E, F | **Set A / Set B** | the two search sets this row clashes — *hidden helper columns*, read by export | — |

- **Tolerance and clearance are opposite functions** — a test is one or the other, never both. If any
  row has **both** Tol and Clr filled, `ExportClashTests` **stops and lists those rows**; nothing is
  written until you clear one of the two on each.
- **Blank Tol and Clr** → no proximity check. **Blank / `None` priority** → no priority.
- **Grouping** is fixed at **15000 mm** on every test (Revizto manages grouping globally), so it has no
  column.
- Distances are stored in the file in **feet**; the macro converts mm for you.
- Re-running `BuildTestList` **preserves** any Tol/Clr/Priority you've edited (matched by Test Name)
  and refreshes the pair list from `Sets`. Delete rows you don't want before exporting.

## Use it

The ready-made starter [`templates/Example.xlsm`](templates/Example.xlsm) already has both modules
pasted in. Or set up your own workbook once: a sheet named **`Sets`** with a header in `A1` and your
search-set names in `A2` downward, then `Alt+F11` ▸ `Insert ▸ Module` and paste
[`src/modClashLite.bas`](src/modClashLite.bas) (and [`src/modImportSets.bas`](src/modImportSets.bas)
for the `.vimsst` name import).

1. **`BuildTestList`** (`Alt+F8`) — fills the `Tests` sheet with every pair + default Tol/Clr/Priority.
2. **Adjust** Tol / Clr / Priority on the `Tests` sheet (drive them from your own matrix), and delete
   any rows you don't want.
3. **`ExportClashTests`** — writes **`Clash Tests.vimctst`** next to the workbook.
4. **Import** the `.vimctst` into Revizto. The set names must match your real Revizto sets.

To edit an **existing** Revizto export instead, run **`ImportClashTests`**, pick the `.vimctst`, edit,
then `ExportClashTests`.

> ⚠️ **Re-importing edits *duplicates*, it doesn't replace.** Every generated test carries a freshly
> minted GUID, and Revizto identifies tests by GUID — so importing an edited set back into the same
> project adds new tests alongside the originals. **Before re-importing, delete the original tests in
> Revizto** (or import into a clean Clash Detection folder).

### Pulling set names from Revizto

Run `ImportSetsFromVimsst` (in `modImportSets`). It shows a file picker — point it at a **`.vimsst`**
search-set export, and it fills the **`Sets`** sheet (column A) with the set names exactly as stored,
so they match on import. Existing names are replaced (you're asked to confirm first). To try it without
your own export, point the picker at the bundled
[`templates/Testing searchsets.vimsst`](templates/Testing%20searchsets.vimsst).

> **OneDrive gotcha:** the macro saves next to the workbook via `ThisWorkbook.Path`. For a workbook in
> OneDrive/SharePoint that path can be an `https://…` URL that file I/O can't use (run-time error 52),
> and Revizto's own exports to a synced folder can land as 0-byte files. Run from a **local, non-synced
> folder** (e.g. `C:\Temp\`) and copy the result back.

## Repository layout

```
src/modClashLite.bas                  BuildTestList / ImportClashTests / ExportClashTests
src/modImportSets.bas                 optional: import set names from a Revizto .vimsst export
docs/vimctst-format.md                reverse-engineered .vimctst format notes
templates/Example.xlsm                ready-to-use starter workbook (both modules pasted in)
templates/Testing searchsets.vimsst   example .vimsst export to try the name import with
tools/sync-xlsm.ps1                    maintainer tool: re-inject the src/*.bas into Example.xlsm
```

### For maintainers — keeping `Example.xlsm` in sync

The `.bas` files are the single source of truth; `Example.xlsm` holds its own embedded copy of each
module. After editing a `.bas`, run [`tools/sync-xlsm.ps1`](tools/sync-xlsm.ps1) to re-import the
modules into the starter (needs Excel + a one-time *Trust access to the VBA project object model*
setting; the workbook must be closed). Point it at another workbook with `-Workbook <path>`.

## Status

Working. Generate, import/round-trip and the `.vimsst` name-import all work; per-test tolerance,
clearance and priority are decoded and written from the `Tests` sheet. A clean starter workbook with
example-only data ships with the repo. **Revizto** (format) and **Architectus** (employer/IP) have both
signed off on publishing.

Decoded but not surfaced in this lite tool: **grouping** (fixed at 15000 mm here), **stamps**, and
**directional clearance** (separate vertical/horizontal distances in one test). Open an issue if you'd
find any of them useful.

## License

[MIT](LICENSE).
