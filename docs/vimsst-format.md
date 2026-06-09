# `.vimsst` — Revizto search-set export (reverse-engineered)

A `.vimsst` is a binary **protobuf** stream, header `MarkerSearchSets`. Same family as
[`.vimctst`](vimctst-format.md). It is a flat list of **field 4** records — folders and search sets,
told apart by `f6`.

## Records

```
header (f1="MarkerSearchSets", f2=1, f3=0)
[ field 4 ] × N        ← folders and sets, in tree order
```

| Field | Folder (`f6=2`) | Search set (`f6=1`) |
|---|---|---|
| `f1` | GUID (36-char ASCII string) | GUID (36-char ASCII string) |
| `f2`,`f3` | timestamps (.NET ticks) | timestamps |
| `f4` | parent-folder GUID (omitted at top level) | parent-folder GUID |
| `f5` | 0 | 0 |
| `f6` | **2** | **1** |
| `f7` | folder name | set name |
| `f8` | — | **filter tree** |

> GUIDs here are **36-char ASCII strings** (unlike the 16 raw bytes used inside `.vimctst`).

## Filter tree (`f8`)

A set's filter is a sequence of condition entries with operator tokens between them:

```
f8 = { f1 = group-op, <entries…>, f3 = trailing-op }
```

- **operator token:** `{ f1 = N }` — boolean / grouping tokens (AND / OR / open-group / close…).
- **condition:** `{ f1=0, f2{ f2{ <leaf> } } }`, where a leaf is:
  ```
  f1 = <kind>
  f2 = { ... = <property group> }     e.g. "Element" / "Other"
  f3 = { ... = <property name>  }     e.g. "Category", "Family and Type"
  f4 = { f1 = <op>, f2 = <value> }    e.g. 5 (equals), "Generic Models"
  f5 = <value type>
  ```

**Conditions are plain strings — no project-specific IDs.** Category / name / parameter conditions
reference properties and values by their *names*, so they are portable and generable from scratch.
The one exception is a condition that **references another set** (e.g. a model search set): that
carries the referenced set's GUID.

A **blank** filter is simply `f8 = { f1=0, f3=0 }` (no conditions) — a valid empty set the user can
fill in.

## Generating sets

Search sets are rule-based filters, not static element lists, so they can be generated:

1. Emit folder records (fabricated GUIDs) to hold the sets.
2. For each set name, emit a set record: fabricated GUID, parent-folder GUID, the name, and a filter
   (blank by default — the user defines the real conditions in Revizto).
3. Write `header + folders + sets`.

> Sets **must** sit inside a folder to work in Revizto: a set's `f4` must reference a folder GUID
> that is present in the same file, otherwise it imports to the root.
