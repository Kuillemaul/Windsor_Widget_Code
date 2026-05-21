YU workbook item-number update patch

What changed
------------
When YU Order Review exports visible rows, it now also writes the canonical Widget item number back into the source YU workbook's visible Sheet1 item-number column (column B by default).

Example:
- Source workbook row 2136 showed: MTS36 BLACK
- Widget / MYOB item is:       MTS36BLACK
- Export still matches row 2136 correctly
- Source workbook Sheet1!B2136 is updated to: MTS36BLACK
- Exported PO also shows: MTS36BLACK

It also updates the in-memory SQL review data for the current session so the candidate/current row display reflects the new item number without waiting for the next workbook import.

Files to replace
----------------
Replace these files in your Windsor Widget source folder:

- main_patched_status_yu.py
- yu_order_workflow.py
- yu_order_review_export_test_window.py

Test
----
1. Close the YU workbook in Excel.
2. Run the app from source.
3. Create/export a YU order for an item where the workbook still has the spaced number.
4. Open the source YU workbook and check the source row, e.g. Sheet1 row 2136 column B.
5. It should now show the no-space item number.

If the workbook is open in Excel, the app should stop and warn that it could not update the workbook.
