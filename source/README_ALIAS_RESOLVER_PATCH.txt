Windsor Widget alias / space-removal resolver patch

Files included:
- main_patched_status_yu.py
- yu_order_workflow.py
- yu_order_review_export_test_window.py

Policy implemented:
- Exact item match is checked first.
- FASTENERS items, identified from Custom List 1 / item group fields, are exact-match only.
- Non-FASTENERS item numbers are resolved using a canonical no-whitespace key.
  Example: ABC 123 and ABC123 both resolve to ABC123.
- Approved duplicate/collision groups can use dbo.item_merge_approvals if present.
- Sales, stock, cover orders, normal order import, item summary, customer summary, order analysis,
  build container / on order checks, and YU order entry/review are wired to the resolver.

Important:
- This patch was syntax-compiled only. It was not connected to the live WindsorWidget SQL Server here.
- Back up the database before installing.
- Test with one or two known items first:
  1. one normal item with old spaces and new no-space format,
  2. one FASTENERS item that must preserve spaces,
  3. one approved collision item.

Suggested SQL test:
- Import or search a normal item as ABC 123 and ABC123; both should show the same Item Summary.
- Search a FASTENERS item using the exact MYOB item number; it should work.
- Search the same FASTENERS item with spaces removed; it should not auto-resolve unless that exact item exists.
