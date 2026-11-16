# Accounting Safety Audit — November 17, 2026

## Change in this build
- Added a posted-journal protection check to the shared module UI.
- When a source record already has a matching `JournalEntries` row with `Status = Posted`, Edit and Deactivate are blocked.
- Protection covers `Invoices`, `Bills`, `CustomerPayments`, `VendorPayments`, `InvoiceLines`, and `BillLines` using their source IDs.
- The message directs users to use a correcting/reversal transaction rather than changing a posted source record.

## Validation performed
- `node --check modules/module-app.js` — passed.
- `node --check modules/accounting-engine.js` — passed.
- Confirmed the six module configurations use the expected sheet/key names in this protection map.

## Scope and limitation
- This is a client-side guard in the existing module UI, not a server-side authorization boundary. Anyone able to modify the client or write directly to the GitHub data repository can bypass it.
- This does not implement reversal workflows. It prevents the current UI from silently changing/deactivating a source record while leaving its posted journal behind.
- Not tested against the user's live GitHub repository. Test first against a backup/copy of the private database.
