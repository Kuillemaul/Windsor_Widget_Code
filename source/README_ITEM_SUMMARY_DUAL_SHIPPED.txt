Windsor Widget - Item Summary dual shipped container patch

Files included:
- main_patched_status_yu.py  (changed)
- yu_order_workflow.py       (included unchanged for safe bundle replacement)
- yu_order_review_export_test_window.py (included unchanged for safe bundle replacement)

What changed:
- Item Summary now groups shipped/on-water orders by arrival date from the orders table.
- The Shipped box shows the earliest shipped/on-water arrival for the selected item.
- If a second shipped/on-water arrival exists, the Next Container box shows that second arrival instead of the normal saved/open next container.
- When Next Container is showing a shipped/on-water arrival, the Next quantity and ETA boxes are highlighted green.
- If no second shipped/on-water arrival exists, Next Container behaves as before.
- Extra shipped/on-water arrivals after the first two are still counted in ordering calculations when they fall inside the calculation horizon, even though only the first two can be displayed.

Test using the supplied orders.csv:
- Items such as MTS36 BLACK, MTS36 CHAR, MT36 WHITE, SA3392 IBC B have multiple arrival dates.
- After importing orders.csv, Item Summary should show the earliest arrival in Shipped and the second arrival in Next.
- The Next box should be green when it is showing a shipped/on-water arrival.

Compile check:
python -m py_compile main_patched_status_yu.py yu_order_workflow.py yu_order_review_export_test_window.py

Run from source:
python main_patched_status_yu.py
