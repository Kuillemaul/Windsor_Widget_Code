# Windsor Widget Import Progress + Fast Canonical Lookup Patch

## Purpose

This patch does two things:

1. Adds a visible progress bar when importing data.
2. Removes the expensive runtime item-number clean-key lookup path from normal app lookups now that imported data is stored in the canonical no-space item-number format.

## User-facing changes

### Import progress bar

The app now shows a progress dialog when importing:

- Sales
- Orders
- Stock
- Cover orders

The progress dialog shows the current import stage, such as:

- reading file
- checking rows
- clearing existing rows
- saving rows
- committing import
- refreshing displays

This stops the app from looking like it has frozen during larger imports.

### Faster item lookups

Normal item-number lookups now use the canonical item number directly instead of repeatedly running space-removal comparisons across large tables.

Example:

```text
Old user input: MTS36 BLACK
Canonical item: MTS36BLACK
```

The app still accepts the spaced version as input, but reporting and lookups are designed to use the stored canonical value.

## Technical notes

### Changed file

```text
main_patched_status_yu.py
```

The included YU files are bundled unchanged so the source folder can be replaced consistently.

### Version

```text
APP_VERSION = 1.2.2
```

### What was removed/avoided

The normal item lookup path no longer relies on recursive clean-key SQL checks such as repeated `REPLACE/LTRIM/RTRIM` matching against large sales/order/stock tables.

The resolver remains for manual input and imports, but it now canonicalises normal imported item numbers directly.

### FASTENERS

Known FASTENERS item numbers are still preserved when they exactly match the item master.

## Install

Replace the files in your source folder with the patched files.

Then run:

```powershell
python -m py_compile main_patched_status_yu.py yu_order_workflow.py yu_order_review_export_test_window.py
python main_patched_status_yu.py
```

## Test checklist

1. Import sales and confirm the progress bar appears.
2. Import orders and confirm the progress bar appears.
3. Import stock and confirm the progress bar appears.
4. Import cover orders and confirm the progress bar appears.
5. Search Item Summary with a no-space item number.
6. Search Item Summary with an older spaced item number.
7. Confirm both searches resolve to the correct item.
8. Check previous sales speed on a high-sales item.
9. Check Shipped and Next Container still display correctly.
