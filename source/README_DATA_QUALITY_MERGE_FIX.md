# Windsor Widget v1.6.3

## Safe Duplicate Item Merge

This release fixes the duplicate item-master repair introduced in v1.6.2.

### Problem fixed

Exact duplicate item numbers could not be distinguished because the merge used the item-number text as the row identifier. If two physical rows both contained the same item number, the Widget could report success without merging their fields.

### New behaviour

- Adds a stable `item_master_id` identity column to `dbo.items`.
- Shows each duplicate as a separate database row with its ID, supplier and populated-field count.
- Copies every blank writable field into the selected canonical record, including supplier data.
- Keeps the selected canonical value when both rows contain conflicting nonblank data.
- Updates historical item-number references when the duplicate uses a different alias spelling.
- Deletes only the physical duplicate rows selected for merging.
- Verifies that exactly one canonical item-master row remains before committing.
- Reports rows deleted, references updated, fields copied and any unresolved conflicts.

### Installation

Replace `main_patched_status_yu.py`, rebuild the application, then restart Windsor Widget.

### Important

The fix cannot recover data that was already deleted by v1.6.2. Restore that field from a database backup, the original item-master source, or enter it manually through Data Quality.
