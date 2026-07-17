# Windsor Widget v1.6.7

## MYOB Purchase Order Export Fix

All MYOB purchase-order exports now omit these columns completely:

- `Currency Code`
- `Exchange Rate`

This applies to:

- **Export MYOB PO** from the YU validation screen
- **Export MYOB Container PO** from Build Container

The fields are no longer included in the TXT header or data rows, so they will not appear in AccountRight field matching.

## Updated Files

```text
main_patched_status_yu.py
yu_order_review_export_test_window.py
```

Replace both files and rebuild Windsor Widget as normal.
