YU canonical export fix

Reason:
- The YU Order Review screen shows the Widget/order CSV item number in the main table.
- It still matches the old spaced source row in the YU workbook using the space-insensitive resolver.
- The current export function only receives source row + qty, so it copied the source workbook item number back into the exported file.

Fix:
- export_yuchang_po_compact_by_rows now accepts source row + qty + canonical item number.
- Export Visible now passes the canonical Widget item number from the review row.
- During export, the matched row's item column is overwritten with the canonical no-space item number while all other formatting and source data is preserved.

Test:
1. Open YU Order Entry.
2. Add MTS36BLACK, qty 100.
3. Open YU Order Review. It should resolve to the Sheet1 source row containing MTS36 BLACK.
4. Export Visible.
5. Open the exported workbook. The exported item column should show MTS36BLACK, not MTS36 BLACK.
