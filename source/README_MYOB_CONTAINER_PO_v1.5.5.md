# Windsor Widget v1.5.5

## MYOB Container PO Comment Order Fix

- `\ON` order headings remain above each original order group.
- Each item is now exported before its matching `\COMMENT` row.
- The comment therefore appears directly below the item in MYOB.
- General container notes are still added at the end of the PO.

### Updated file

- `main_patched_status_yu.py`

Replace the file and rebuild Windsor Widget as normal.
