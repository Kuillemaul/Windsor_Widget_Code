# Windsor Widget v1.7.3

## Supplier Container Export — Comments Column Removed

The supplier-facing Excel export created from **Build Container** no longer includes the `Comments` column.

The exported columns are now:

```text
Order Number
Item Number
Description
Qty
Cartons
```

## Unchanged Behaviour

This update does not remove or alter:

- Internal container-line comments stored in Windsor Widget
- General container notes
- MYOB `\\COMMENT` rows in the MYOB purchase-order export
- The MYOB manual-change worksheet

Only the Comments column in the container workbook intended for emailing to the supplier has been removed.

## Updated File

```text
main_patched_status_yu.py
```

## Installation

1. Close Windsor Widget.
2. Back up the current source file.
3. Replace `main_patched_status_yu.py`.
4. Rebuild or restart the application as normal.

No database changes are required.
