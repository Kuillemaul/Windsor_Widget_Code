# Windsor Widget v1.7.5

## YU Order Validation — Mapping Sheet Fix

This update clarifies and repairs the relationship between the Yuchang workbook sheets.

## Diagnosis

The workbook contains separate sheets with different purposes:

- `Sheet1` — the supplier order form. Supplier item rows are stored here.
- `ITEM_Yuchang_Match` — the Windsor item-master lookup list.
- `Match_Review` — historical matching suggestions and review data.

The row numbers are **not shared between these sheets**.

For example, row 48 on `ITEM_Yuchang_Match` is simply the 48th item-master entry. It does not identify supplier row 48 on `Sheet1`.

The permanent supplier-line mapping belongs in:

```text
Sheet1!A<supplier row>
```

## Changes

- Detects the YU supplier order form from its `Item`, `Size`, and `Colour` headers.
- Prefers `Sheet1` when it is the valid order-form sheet.
- Never treats `ITEM_Yuchang_Match` as a row-aligned supplier sheet.
- Writes the selected Windsor item number into the exact order-form mapping cell.
- Reopens the workbook and verifies the saved cell before reporting success.
- Error messages now identify exact cells such as:

```text
Sheet1!A48
```

- The validation window title now displays:

```text
YU Order Review — Validation 1.7.5
```

This makes it clear that the rebuilt EXE contains the current validation module.

## Verified Behaviour

Using the supplied Yuchang workbook:

- The order form was correctly detected as `Sheet1`.
- Supplier row 48 was recognised as an item-detail row.
- A test mapping was written to `Sheet1!A48`.
- `ITEM_Yuchang_Match!A48` remained unchanged.
- The saved mapping was reopened and verified.
- A compact test supplier-order export resolved the item from `Sheet1!A48`.
- All included Python files pass `py_compile`.

## Installation

Replace these files using the exact filenames shown:

```text
main_patched_status_yu.py
yu_order_review_export_test_window.py
yu_order_workflow.py
Windsor Widget.spec
```

Do not leave numbered download names such as:

```text
yu_order_review_export_test_window(54).py
```

PyInstaller imports the file named exactly:

```text
yu_order_review_export_test_window.py
```

## Build Command

Run from the source folder:

```powershell
python -m PyInstaller --noconfirm --clean "Windsor Widget.spec"
```

After rebuilding, open order validation and confirm that the window title includes:

```text
Validation 1.7.5
```

No database changes are required.
