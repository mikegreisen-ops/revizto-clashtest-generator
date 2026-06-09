# templates/

This folder will hold a **scrubbed, ready-to-use starter workbook** — the public artifact users
download to begin.

It must be a *clean copy* of the generator workbook with **all project-specific data removed**:

- example/placeholder elements only (no real project element list)
- cleared override grids (clearance / priority / ignore)
- no real search-set names, GUIDs, or model-set references
- no author email or project identifiers embedded in any macro template constants

The live working workbook is **git-ignored** and must never be committed here. When the starter
workbook is ready, force-add it explicitly:

```
git add -f templates/clash-generator-starter.xlsm
```
