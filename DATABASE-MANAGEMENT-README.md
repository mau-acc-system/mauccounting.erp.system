# Database Management module

- Open from Administration > Setup > Database Management.
- Administrator-only screen.
- Reads every entity/sheet from `data/manifest.json` through the existing GitHub `.dat` storage API.
- Supports sheet tabs, search, row selection, add, edit, delete, and CSV export.
- Saves changes to the private database repository using the existing GitHub commit API.
- Edit/delete are blocked when a matching posted journal source reference is detected. This is a guardrail, not a substitute for full reversal workflows.
- `excel.js` exposes `getSheetHeaders(sheet)` for schema-aware table columns.

Upload the contents of this package to the root of the public `mauccounting.erp.system` repository. Do not upload the private database contents or `license.tk` to the public repository if your current deployment keeps them private; the existing architecture references `license.tk` in the public frontend and should be reviewed separately.
