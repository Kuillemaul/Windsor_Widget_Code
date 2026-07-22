# Windsor Widget v1.7.1

## Live Refresh View-Position Fix

This patch stops the multi-user live refresh from dragging the selected row back into the centre of the table.

### Fixed Screens

- To Order
- On Order
- Build Container

### New Behaviour

- The selected row is retained after a refresh when the same record still exists.
- The selected column is retained.
- Vertical scroll position is retained.
- Horizontal scroll position is retained.
- A selected row may remain off-screen while the user reviews another part of the sheet.
- Automatic and manual refreshes use the same view-preservation behaviour.
- Build Container also preserves the view and selection inside the current container lines table.

### Database Changes

None.

### Updated File

```text
main_patched_status_yu.py
```

### Installation

1. Close Windsor Widget.
2. Back up the current source file.
3. Replace `main_patched_status_yu.py` with the patched file.
4. Rebuild the application as normal.

### Verification

- Python byte-code compilation passed.
- A full PySide6 GUI test was not performed in the patch environment because PySide6 is not installed there.
