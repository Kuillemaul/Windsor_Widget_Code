# Windsor Widget Sales Import Customer Sync Patch

## Purpose

When sales are imported from MYOB, any customer found in the sales file that is not already in the Windsor Widget customer list is now added automatically.

This means the customer can be searched from the Customer Summary screen even before a customer Excel workbook has been matched.

## User-facing behaviour

- Import Sales now checks the customer list during import.
- New customers from the import file are added to the customer list automatically.
- The Sales Import completion message now shows how many new customers were added.
- Customer Summary search/autocomplete also includes customer names from sales history, so existing sales customers remain searchable even if they were not previously in the customer list.

## Notes

New customers are added with no matched Excel file. The customer workbook can still be matched later from the Customer Summary screen.

## Version

APP_VERSION = 1.2.3
