# Revizto Clash-Test & Search-Set Generator

Generate Revizto **clash tests** (`.vimctst`) and **search sets** (`.vimsst`) in bulk from a single
Excel workbook — instead of building and naming hundreds of them by hand in Revizto.

A BIM coordination project can need *thousands* of clash tests (every discipline pair, at the right
priority and clearance). Creating and maintaining those by hand is slow and error-prone. This tool
treats one Excel workbook as the source of truth and emits ready-to-import Revizto files from it.

> ⚠️ **Unofficial.** These file formats are **reverse-engineered**, not documented or supported by
> Revizto. This project is **not affiliated with or endorsed by Revizto**. Formats can change between
> Revizto versions. **Always keep backups and work on copies.** Use at your own risk.

---

## How it works

The Revizto files are length-delimited **protobuf** with no file-wide checksum, so records can be
generated and re-imported. Two facts make bulk generation practical:

1. **Clash tests reference their search sets by *name*, not GUID.** On import, Revizto re-matches a
   test's two sides to existing project sets by the set **name string** — the stored GUID and folder
   path are cosmetic. So generated tests need only the *names* of the user's sets (no GUID harvest).
2. **Search-set filter conditions are plain strings.** Category / parameter / name conditions carry
   no project-specific IDs, so simple sets can be generated from scratch.

This means the whole pipeline can run from the workbook alone.

## The three workflows

| Workflow | You have… | The tool… | Macro |
|---|---|---|---|
| **C — Tests only** | search sets already in Revizto, matching your list | generates the clash tests | `modClashGen` |
| **B — Sets first** | search sets built in Revizto | imports their names into the workbook, then you generate tests | `modImportSets` → `modClashGen` |
| **A — From scratch** | just a list of elements | generates blank search sets to import, *and* the tests | `modGenSets` → `modClashGen` |

See [`docs/workflows.md`](docs/workflows.md) for the detail.

## Repository layout

```
src/        VBA modules — paste into the workbook (Alt+F11 ▸ Insert ▸ Module)
docs/       reverse-engineered file-format notes + workflow guide
templates/  (todo) a scrubbed, ready-to-use starter workbook
```

## Install & use

1. Open your generator workbook in Excel.
2. `Alt+F11` ▸ `Insert ▸ Module`, paste the contents of a file from `src/`. Repeat per module.
3. Run the relevant macro (e.g. `GenerateClashTests`). The output file lands beside the workbook.
4. Import the output into Revizto.

> **OneDrive gotcha:** the macros write next to the workbook via `ThisWorkbook.Path`. For a workbook
> stored in OneDrive/SharePoint that path can be an `https://…` URL that file I/O can't use
> (run-time error 52). Run from a **local, non-synced folder** (e.g. `C:\Temp\clashgen\`) and copy
> the result back.

## File formats

- [`docs/vimctst-format.md`](docs/vimctst-format.md) — clash-test set format
- [`docs/vimsst-format.md`](docs/vimsst-format.md) — search-set format

## Status

Private / work in progress. Workflows B and C are working; workflow A's set generator is being
finalised. **Before any public release:** confirm Revizto are comfortable with it, clear any
employer/IP questions, and scrub all project-specific data from the modules and template.

## License

[MIT](LICENSE) — pending IP review (see Status).
