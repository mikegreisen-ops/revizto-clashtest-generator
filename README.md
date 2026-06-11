# Revizto Clash-Test Generator

Build a Revizto **clash test for every pairing of your search sets** — and set each test's tolerance,
clearance, priority and stamp — from **one Excel workbook**, instead of creating and configuring
hundreds of tests by hand in Revizto. It can also read an existing clash-test export **back into Excel**
so you can bulk-edit settings and re-export.

**Made for BIM coordinators — no coding needed.** You fill in a spreadsheet, click a button, and import
the result into Revizto.

> ⚠️ **Unofficial community tool.** It reads/writes Revizto's clash-test format (`.vimctst`), which is
> **reverse-engineered** — Revizto don't document or support it, and it can change between versions.
> **Revizto have kindly approved sharing this**, but it isn't a Revizto product and they don't maintain
> or back it — so **always keep backups and work on copies**. Use at your own risk.

---

## Quick start

### 1. Get the workbook (you do **not** need a GitHub account)

Just one file:

1. Open [`templates/Example.xlsm`](templates/Example.xlsm) above, then click **Download raw file** (the
   ⤓ download icon near the top-right of the file view). That single workbook has everything built in.
   *(Or use the green **Code ▸ Download ZIP** button to grab the whole project — but you only need the
   one `.xlsm`.)*
2. **Unblock it.** Windows blocks macros in files downloaded from the internet. Right-click the
   downloaded `.xlsm` → **Properties** → tick **Unblock** (bottom of the General tab) → **OK**.
3. Open it and click **Enable Content / Enable Macros** when Excel asks.

> Skipping the Unblock step is the #1 "it doesn't work": Excel shows a red *"macros have been blocked"*
> banner and the buttons won't run. It's not broken — just blocked. Unblock, reopen, enable macros.

### 2. Add your search-set names

The workbook opens on a sheet called **`Sets`**. Put your Revizto search-set names in **column A** (one
per row, from row 2). Two ways:

- **Recommended — import them** so there are no typos: run the **`ImportSetsFromVimsst`** macro
  (`Alt+F8` to see the macro list), and point the file picker at a **`.vimsst`** search-set export from
  Revizto. It fills column A for you, exactly as named. *(To try it first, point it at the bundled
  [`templates/Testing searchsets.vimsst`](templates/Testing%20searchsets.vimsst).)*
- **Or type / paste** the names in yourself — but they must match your Revizto set names **exactly**.

### 3. Build → adjust → export

1. Run **`BuildTestList`** (`Alt+F8` → pick it → Run). It fills a **`Tests`** sheet with every pairing
   of your sets, each with sensible default settings.
2. **Adjust** the Tolerance / Clearance / Priority / Stamp columns to suit (or delete any rows you don't
   want). You can drive these straight from your own coordination matrix.
3. Run **`ExportClashTests`**. It writes **`Clash Tests.vimctst`** next to the workbook.
4. **Import** that file into Revizto (Clash Detection ▸ import). Done — all your tests, named and
   configured.

> "Every pairing" means every set against every other set, **plus each set against itself** — so 10
> sets makes 55 tests. (You can delete the self-pairs or any others you don't need before exporting.)

---

## The three buttons

| Macro (`Alt+F8`) | What it does |
|---|---|
| **`BuildTestList`** | `Sets` list → `Tests` sheet (every pairing, with editable settings) |
| **`ExportClashTests`** | `Tests` sheet → `Clash Tests.vimctst` (the file you import into Revizto) |
| **`ImportClashTests`** | an existing `.vimctst` → `Tests` sheet (to edit tests you already have) |
| **`ImportSetsFromVimsst`** | a `.vimsst` → fills the `Sets` list with your set names |

## The `Tests` sheet

`BuildTestList` and `ImportClashTests` both fill this sheet; `ExportClashTests` reads it.

| Column | Meaning | Default |
|---|---|---|
| **Test Name** | `"<Set A> vs <Set B>"` — rename it however you like | — |
| **Tolerance (mm)** | how far things must overlap to count as a clash | `25` |
| **Clearance (mm)** | a required gap — clashes when things are *closer* than this | *(blank)* |
| **Priority** | `Trivial` / `Minor` / `Major` / `Critical` / `Blocker` | `Minor` |
| **Stamp** | a Revizto stamp code, e.g. `0AR` (blank = no stamp) | *(blank)* |

(Two more columns, **Set A** and **Set B**, are hidden far to the right — the tool needs them, you
don't.)

Good to know:

- **Tolerance and Clearance are opposites** — a test uses one or the other, never both. If you fill in
  both on a row, `ExportClashTests` **stops and lists those rows** so you can fix them; nothing is
  written until you do.
- **Stamp codes are CASE-SENSITIVE** — `0AR` is not `0ar`. The code must match a stamp that already
  exists in your Revizto project exactly, or no stamp is applied (silently).
- **Re-running `BuildTestList` keeps your edits** — it remembers the settings you changed (by test name)
  and just refreshes the list of pairings from `Sets`.
- Grouping is set to a fixed **15 m** on every test (Revizto manages grouping globally anyway).

## Editing tests you already have (round-trip)

Already built tests in Revizto and want to bulk-change their settings? Export them from Revizto, then:

1. Run **`ImportClashTests`** and pick the `.vimctst` — it loads every test onto the `Tests` sheet.
2. Edit Tolerance / Clearance / Priority / Stamp.
3. Run **`ExportClashTests`** to write the updated file.

> ⚠️ **Re-importing makes *new* tests, it doesn't overwrite the old ones.** Revizto tells tests apart by
> a hidden ID, and every export here gets fresh IDs — so importing your edited file alongside the
> originals gives you **duplicates**. Before re-importing, **delete the original tests in Revizto** (or
> import into an empty Clash Detection folder).

## A couple of gotchas

- **OneDrive / SharePoint:** the macro saves the `.vimctst` next to the workbook. If the workbook lives
  in a synced OneDrive/SharePoint folder, that can fail (Excel sees a web address, not a real path —
  run-time error 52), and Revizto's own exports to synced folders sometimes save as empty files. **Work
  from a normal local folder** (e.g. `C:\Temp\`) and copy files back afterwards.
- **Names must match.** A clash test binds to your sets by **name** (and a stamp by its **code**). If a
  name/code doesn't match anything in your Revizto project, that side imports with a warning and a
  "re-select from list" prompt — harmless, but you'll have to fix it. Importing your set names (step 2)
  is the surest way to avoid this.

---

## How it works (for the curious)

When you import a clash test, **Revizto matches each side to one of your project's search sets by the
set's *name*** — the internal IDs and folder paths stored in the file are ignored for matching. So the
tool only needs your set *names*; it makes up the IDs. (Priority and stamps work the same way — matched
by value/code, not by ID.) That's why a wrong name is obvious and harmless rather than destructive: it
just doesn't match, and Revizto asks you to pick the right set.

## For developers

The macros live as plain text in [`src/`](src/) so they can be reviewed and version-controlled; the
`.xlsm` simply carries a pasted copy. To set up your own workbook instead of the starter: add a sheet
named `Sets`, then `Alt+F11` ▸ `Insert ▸ Module` and paste
[`src/modClashLite.bas`](src/modClashLite.bas) (and [`src/modImportSets.bas`](src/modImportSets.bas) for
the set-name import).

```
src/modClashLite.bas                  BuildTestList / ImportClashTests / ExportClashTests
src/modImportSets.bas                 ImportSetsFromVimsst (set names from a .vimsst)
docs/vimctst-format.md                reverse-engineered .vimctst format notes
templates/Example.xlsm                ready-to-use starter workbook (modules pasted in)
templates/Testing searchsets.vimsst   example .vimsst to try the set-name import with
tools/sync-xlsm.ps1                    re-injects src/*.bas into Example.xlsm (keeps them in sync)
```

**Keeping `Example.xlsm` in sync:** the `.bas` files are the source of truth. After editing one, run
[`tools/sync-xlsm.ps1`](tools/sync-xlsm.ps1) to re-import the modules into the starter (needs Excel and a
one-time *Trust access to the VBA project object model* setting; the workbook must be closed). Target
another workbook with `-Workbook <path>`.

## Status

Working. Building tests, the round-trip import/export, and the `.vimsst` name import all work; per-test
tolerance, clearance, priority and stamps are written into the file. The starter workbook ships with no
real project data. **Revizto have approved sharing this.**

## License

[MIT](LICENSE).
