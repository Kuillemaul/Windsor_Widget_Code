# Windsor Widget v1.5.4

## MYOB Container PO Export

This update adds an **Export MYOB Container PO** button to the **Build Container** screen.

### What it creates

Each export creates two files:

1. **MYOB_CONTAINER_PO_[PO NUMBER].txt**
   - Imports one new Yuchang purchase order into AccountRight.
   - Groups the items by their original order number.
   - Uses `\ON` rows for original order headings.
   - Uses `\COMMENT` rows for comments.
   - Leaves Supplier Invoice No. blank.
   - Uses the stored item description and latest YU cost.
   - Falls back to the generic latest purchase cost when no YU-specific cost is stored.

2. **MYOB_MANUAL_CHANGES_[PO NUMBER].xlsx**
   - Lists every original MYOB order that must be changed manually.
   - Shows item number, description and quantity moved to the new container PO.
   - Includes a Done column for checking off completed changes.

### Build Container comments

- A **Comments** column has been added to the container table.
- Comments from **On Order** are carried into the container.
- Double-click a container comment to add or edit it.
- Container notes and the dog-leads note are also exported as `\COMMENT` rows.

### Workflow

1. Build or load the container.
2. Confirm the original order number on each line.
3. Add any required comments.
4. Click **Export MYOB Container PO**.
5. Enter the new MYOB purchase order number and date.
6. Import the generated TXT through:
   - **File → Import/Export Assistant**
   - **Purchases → Item Purchases**
7. Use the generated Excel cheat sheet to manually reduce or remove the listed items from the original MYOB purchase orders.

### Database update

The Widget automatically adds this field when required:

```sql
container_lines.comments
```

Existing container lines are retained and start with a blank comment.

### Installation

Replace:

```text
main_patched_status_yu.py
```

Then rebuild the Windsor Widget application as normal.

### Test file

`MYOB_Test_ON_COMMENT_999999TEST.txt` is included to verify that AccountRight accepts `\ON` and `\COMMENT` rows before using the live container export.
