# Windsor Widget v1.6.2

## Data Quality Scan Control and In-Page Fixes

### Changes

- Data Quality no longer scans automatically when the page is opened.
- Added **Cancel Scan** for stopping a long scan.
- Leaving the Data Quality page also cancels an active scan.
- Existing results are retained when a scan is cancelled.
- Added **Fix Selected**.
- Double-click **Suggested action** to open the relevant repair dialog.
- Double-click another table column to open Item Summary.

### Supported in-page repairs

- Add missing descriptions.
- Enter missing latest purchase costs and dates.
- Assign missing suppliers.
- Enter missing carton quantities.
- Create item-master records referenced by Sales, Stock, Orders, or the YU template.
- Merge duplicate or ambiguous item-master records after selecting the canonical item.
- Assign an item number to a blank item-master record.
- Correct duplicate YU template column-A mappings, with an automatic workbook backup.
- Choose a replacement YU template when the current workbook cannot be scanned.

Potentially destructive changes, such as merging duplicate item records, require confirmation.
After a repair, press **Refresh Checks** to confirm all related issues are cleared.

## Installation

Replace `main_patched_status_yu.py`, then rebuild the application as normal.
