Windsor Widget - YU Export Canonical Item Number Patch

What changed:
- YU export still matches the correct supplier workbook row by source row.
- The exported workbook now writes the Widget/resolved item number into the item number column (default column B) instead of copying the old spaced item number from the YU template row.
- Quantity/date/order-number/export formatting behaviour is otherwise unchanged.

Example:
- YU workbook/template row contains: ABC 123
- Widget/YU order line resolves to: ABC123
- Exported workbook now shows: ABC123

Files included:
- main_patched_status_yu.py
- yu_order_workflow.py
- yu_order_review_export_test_window.py
- README_ALIAS_RESOLVER_PATCH.txt
- README_YU_EXPORT_CANONICAL_ITEM_PATCH.txt

Install/test:
1. Back up the current source folder.
2. Replace the three .py files with these versions.
3. Run:
   python -m py_compile main_patched_status_yu.py yu_order_workflow.py yu_order_review_export_test_window.py
4. Run the app from source and export a YU order using an item that matched a spaced workbook row.
5. Open the exported workbook and confirm the item number column shows the no-space/resolved item number.
