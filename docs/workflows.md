# Workflows

Everything hangs off one workbook. Its `SearchSets` sheet is the source of truth — one row per
element, with at least a **Code** and a **set name** column (the name must match the element's
search set in Revizto). A matrix of element pairs (with priority/clearance) produces the test names.

Pick the workflow that matches what you already have in Revizto.

## C — Tests only ("trust me, the sets already line up")

You've already created the search sets in Revizto and named them to match the workbook.

```
workbook ──(modClashGen)──▶ .vimctst ──▶ import to Revizto
```

`GenerateClashTests` reads the pairs + set names and writes a clash-test file. Because Revizto
matches sets **by name**, that's all it needs — no GUIDs, no set export. A mistyped name imports with
a warning and a "re-select" prompt, so mismatches are obvious and harmless.

## B — Sets first ("pull the real names back in")

You built the sets in Revizto and want the workbook to use their exact names.

```
Revizto sets ──export──▶ .vimsst ──(modImportSets)──▶ workbook ──(modClashGen)──▶ .vimctst
```

`ImportSetsFromVimsst` reads a `.vimsst` and lists every set name (and folder) on an `ImportedSets`
sheet, and flags any workbook set-name that has no match — so you can align names before generating
tests. The most robust path, since the names are guaranteed to exist.

## A — From scratch ("I just have a list")

You have the element list but no sets yet.

```
workbook ──(modGenSets)──▶ .vimsst (blank sets) ──▶ import to Revizto
   then refine each set's filter in Revizto
workbook ──(modClashGen)──▶ .vimctst ──▶ import to Revizto
```

`GenerateSearchSets` emits one folder plus a **blank** search set per name, ready to import. The sets
arrive correctly named and foldered; you define each set's actual filter conditions in Revizto (the
part only you can specify). Then generate the tests as in workflow C.

---

## What still needs doing by hand in Revizto

Generation creates the *tests and sets*; everything else is a normal Revizto pass:

- Define / refine each search set's filter conditions (workflow A).
- Bulk-edit clearances, tolerances, stamps, and view settings via filters & tags.
- Anything project-specific the workbook doesn't model.

The point of the tool is to remove the thousands of manual create-name-pair clicks, not the
coordination judgement.
