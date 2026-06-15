# `.vimctst` — Revizto clash-test set (reverse-engineered)

A `.vimctst` is a binary **protobuf** stream (length-delimited fields), header `RClashTests`. No
file-wide checksum, so records can be generated/edited and re-imported. Don't open it in a text
editor — it will corrupt.

> Decoded against **Revizto 1.17.6** (Windows). The format isn't documented and can change between
> versions; if a later version breaks generation, re-harvest the embedded template from a fresh export.

## Top-level layout

The file is **three parallel arrays**, grouped (not interleaved), joined by GUIDs:

```
header  (f1="RClashTests", f2=1, f3=0)
[ N × field 4 ]   ← test identity
[ N × field 5 ]   ← link / settings
[ N × field 6 ]   ← the clash rule
```

Each test spans one record in each array. They are joined by GUIDs:
`field4.f1 == field5.f1` (test GUID), and `field5.f7 == field6.f1` (rule GUID).

### field 4 — identity
| Field | Meaning |
|---|---|
| `f1` | test GUID (16 **raw** bytes) |
| `f2` | test name (string) |
| `f4`–`f9` | flags (varints) |
| `f10` | `"Ignore"` — present only on ignored tests |
| `f11` | `{ f1 = author email }` |

### field 5 — link / settings
| Field | Meaning |
|---|---|
| `f1` | test GUID (== `field4.f1`) |
| `f2` | result/mid GUID (unique per test) |
| `f7` | rule GUID |
| `f16`+ | report/display settings (mostly defaults) |

### field 6 — the clash rule
| Field | Meaning |
|---|---|
| `f1` | rule GUID (== `field5.f7`) |
| `f2` | **Element A** side |
| `f3` | **Element B** side |
| `f4` | clash **mode + distance** — `f4.f3` = 3× float32 distance **[horizontal, vertical, horizontal]** (**feet**), `f4.f4` = mode (`0` none, `1` clearance, `2` tolerance, `3` directional clearance) |
| `f9` | `f9.f1.f3.f1` = float32 **grouping** distance (feet) |
| `f10` | **priority + stamp** — `f6` = priority (`1` Trivial … `5` Blocker, `0` none); `f1` = flag bits (priority sets `4`, stamp sets `1`); `f4` = stamp code string; `f11` = stamp GUID (16 raw bytes) |
| `f5`–`f8`, `f11` | colour / display / report settings (cloned from template) |

Each **side** (`f2`/`f3`) is:
```
f1    = set GUID (16 raw bytes)
f2[]  = path strings, e.g. "Clash Detection","00_Categories","<folder>","<set name>"
```

## Key facts for generation

- **Sets are matched by name on import.** Revizto re-binds each side to a project search set by the
  **set-name string** (the last path element). The set **GUID and folder path are cosmetic** — a
  generated test can carry a fabricated GUID and any folder path; only the name must match a real
  set in the target project. A name with no match imports with a warning and a "re-select from list"
  prompt (no corruption).
- **Most settings are cloned from a template** and bulk-edited in Revizto afterward — colour, display,
  report flags. But **clearance/tolerance, clash mode (incl. directional), grouping distance, priority
  and stamp are now decoded** and can be written per test (see "Per-test settings" below).
- **GUID byte order:** the 16 raw bytes equal `Guid.ToByteArray()` of the 36-char string form
  (.NET little-endian for the first three components). Freshly minted GUIDs are already in that order.

## Per-test settings (decoded 2026-06-11)

Distances are stored in **feet** — multiply mm by `1/304.8`. Located by single-variable diff-probes:
export two tests differing in exactly one setting, diff the bytes (see `diff_vimctst.ps1`).

| Setting | Bytes | Encoding |
|---|---|---|
| Clearance / tolerance distance | `field6.f4.f3.{f1,f2,f3}` | 3× float32, **feet** — axes **[horizontal, vertical, horizontal]**; all equal for a uniform distance, different for directional clearance |
| Clash mode | `field6.f4.f4` | varint — `0` none, `1` clearance (uniform), `2` tolerance, `3` directional clearance |
| Grouping distance | `field6.f9.f1.f3.f1` | float32, feet |
| Priority | `field6.f10.f6` | varint — `1` Trivial, `2` Minor, `3` Major, `4` Critical, `5` Blocker, `0` none (`f10.f1` flag bit `4` set when present) |
| Stamp | `field6.f10` | `f1` flag bit `1`, `f4`=code string (e.g. `"0AR"`), `f11`=GUID (16 raw bytes) |

**Stamps bind by code, not GUID** — like search sets. Revizto re-matches a test's stamp to the
project's stamp by the **code string** (`f10.f4`); the stamp GUID (`f10.f11`) is cosmetic. Proven by
importing a test with the right code but a garbage GUID (the `STAMP_BADGUID` probe) — the stamp bound
fine. So a generator needs only the code and can fabricate the GUID.

The original sample test carried **no stamp** (`f10` had no `f4`/`f11`), so adding a stamp *inserts*
those sub-fields and grows `f10` — unlike clearance, which is a fixed-length overwrite of existing floats.

## Generating a file

1. Read the element pairs (which set clashes against which) and the test names from the workbook.
2. For each test, mint 3 GUIDs (test, mid, rule) and emit a `field4` + `field5` + `field6` record,
   cloning the settings tail from a template and writing each side's **set name**.
3. Write `header + [all field4] + [all field5] + [all field6]`.
4. Verify: parses to EOF; equal record counts; `f4.f1==f5.f1`; `f5.f7==f6.f1`; names unique.
