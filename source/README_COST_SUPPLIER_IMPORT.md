# Windsor Widget v1.6.5

## Cost, Description and Supplier Import

The MYOB purchase-history importer now updates the item's last purchased supplier as well as its latest cost.

### Changes

- Renamed the action to **Update Costs, Descriptions & Suppliers**.
- Reads `Co./Last Name` or `Supplier` from the MYOB Item Purchases export.
- Assigns the supplier from the newest valid positive purchase row for each matched item.
- Updates the supplier even when the cost is already current, provided the source purchase date is not older than the stored latest-cost date.
- Full Refresh uses the newest cost and supplier found in the selected complete-history file.
- Adds imported supplier names to the Widget supplier list.
- Continues to fill descriptions only when both description fields are blank.
- Continues to protect newer stored costs and suppliers during normal incremental imports.
- Keeps Yuchang-specific costs current for the MYOB YU PO export.

### Recommended MYOB fields

```text
Item Number
Description
Price
Date
Co./Last Name
```

Without `Co./Last Name` or `Supplier`, costs and descriptions still import, but supplier assignments are not changed.

### Installation

Replace:

```text
main_patched_status_yu.py
```

Then rebuild Windsor Widget as normal.
