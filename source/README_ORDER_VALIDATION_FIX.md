# Windsor Widget v1.7.4

## Order Validation Fix

This release repairs the **Validate and Export** workflow for Yuchang orders.

## Root Causes

Three separate issues could prevent the order-validation window from opening:

1. YU draft, temporary and export folders were based on `__file__`.
   In a PyInstaller one-file build, this can point into PyInstaller's temporary
   extraction directory rather than a stable writable location.

2. Errors raised while creating the validation CSV or opening the review window
   were not caught by the order-entry dialog. The button could therefore appear
   to do nothing.

3. The YU review window still required the legacy `yu_test_*` validation tables.
   Current validation uses Column A of the Yuchang workbook as the permanent item
   mapping, so those historical tables should not block normal validation.

## Changes

### Stable YU data folders

YU working files now use a stable per-user location:

```text
%LOCALAPPDATA%\WindsorWidget\YU\drafts
%LOCALAPPDATA%\WindsorWidget\YU\temp
%LOCALAPPDATA%\WindsorWidget\YU\exports
```

The application falls back to `%APPDATA%` and then the system temporary folder
when required.

### Workbook location

The application now looks for the editable workbook beside the installed EXE:

```text
yuchang_order_form_Widget.xlsx
```

It also retains the existing saved-path and development-folder checks.

The workbook should remain external to the EXE because confirmed item mappings
are written permanently into Column A.

### Visible validation errors

**Validate and Export** now catches integration failures and displays the exact
error. A diagnostic file is also written to:

```text
%LOCALAPPDATA%\WindsorWidget\YU\order_validation_error.log
```

### Workbook-only validation mode

The legacy SQL validation tables are now optional.

When they are unavailable:

- Existing Column A item mappings still resolve normally.
- Orders can still be reviewed and exported.
- New mappings can still be confirmed against a current workbook row.
- Historical suggestions and candidate matches are simply unavailable.
- The review window displays that it is using workbook-only validation mode.

### PyInstaller packaging

The spec file now explicitly includes:

```text
yu_order_workflow
yu_order_review_export_test_window
```

This prevents PyInstaller from omitting the validation modules.

## Installation

Replace these four files in the v1 source folder:

```text
main_patched_status_yu.py
yu_order_workflow.py
yu_order_review_export_test_window.py
Windsor Widget.spec
```

Keep the editable workbook beside the EXE or select it when prompted:

```text
yuchang_order_form_Widget.xlsx
```

## Build Command

Run from the source folder with the virtual environment active:

```powershell
python -m PyInstaller --noconfirm --clean "Windsor Widget.spec"
```

The rebuilt executable will be created under:

```text
dist\
```

## Verification

Completed checks:

- All three Python files pass `py_compile`.
- The supplied Yuchang workbook scans successfully.
- A test order line resolves from Column A.
- A test supplier workbook export completes successfully.
- Workbook-only validation resolves mapped items with all legacy validation
  tables unavailable.

No database changes are required.

A full live SQL Server and packaged GUI session was not available in the patch
environment, so the rebuilt application should still be tested with one small
order before normal use.
