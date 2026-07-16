# Windsor Widget v1.6.1

## Data Quality Progress Bar

Adds a live progress bar to the Data Quality Centre.

### Changes

- Shows the current scan percentage and task.
- Reports progress while checking:
  - item-master duplicates
  - missing descriptions, costs, suppliers, and carton quantities
  - sales, stock, order, and container references
  - the live YU template
  - final result preparation
- Keeps the interface responsive during longer scans.
- Disables **Refresh Checks** while a scan is running to prevent duplicate scans.
- Shows **Complete** with the number of issues found, or **Scan failed** if an error occurs.

## Installation

Replace `main_patched_status_yu.py`, then rebuild the application as normal.
