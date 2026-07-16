# Windsor Widget v1.6.0

## Data Quality Centre

A new **Data Quality** page scans the current Widget database and Yuchang template for:

- Sales, stock, order, To Order, On Order, or container items missing from the item master
- Missing item descriptions
- Missing latest purchase costs
- Missing suppliers
- Duplicate canonical item numbers
- Recently sold items with no carton quantity
- Duplicate or unknown Windsor item numbers in YU template column A

Issues are grouped by severity and category. They can be searched, filtered, exported to CSV, or opened directly in Item Summary by double-clicking the row.

The Home dashboard's data-quality button now opens the full Data Quality Centre.

## Watched Import Folder

The **Update Data** page now includes a **Watched Folder** tab.

Choose a root folder and the Widget creates:

```text
Sales\
Stock\
Orders\ToDo\
Orders\Purchases\
Cover Orders\
Item Costs\
Processed\
Failed\
```

Place each MYOB export in the matching folder. The Widget scans every 30 seconds and shows recognised files in an import queue.

Features:

- Manual **Import Selected** and **Import All Ready**
- Optional automatic importing
- SHA-256 duplicate-file detection
- Shared SQL import history across Widget users
- SQL application locking to prevent two users importing the same file together
- Optional movement of successful files into dated `Processed` folders
- Failed imports remain available for retry
- Older sales, stock, cover-order, and open-order snapshots are marked **Superseded** when a newer file is available
- Item-cost files are processed incrementally from oldest to newest
- Open orders are imported only when both the To Do List and Item Purchases files are present

Automatic importing is disabled by default. Enable it only after the folder workflow has been tested with your normal MYOB exports.

## Database Change

The application automatically creates:

```sql
watched_import_history
```

This table stores file fingerprints, type, status, processing time, result message, and archived path. No existing business tables are removed or rebuilt.

## Installation

Replace:

```text
main_patched_status_yu.py
```

Then rebuild the Windsor Widget application as normal.

## Verification Performed

- Python byte-code compilation passed
- Full source AST parsing passed
- No duplicate `MainWindow` methods detected
- Watched-folder creation and file classification tested
- SHA-256 fingerprint caching tested
- SQLite import-history insert/update tests passed
- Data-quality matching and issue-generation tests passed with simulated item, sales, and stock data

A full PySide6 GUI launch and live SQL Server import were not available in the patch environment. Test with copied MYOB exports before enabling automatic imports.
