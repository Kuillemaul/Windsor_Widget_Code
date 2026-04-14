Repo-aware builder
==================

What this does
--------------
This builder is designed for the real repository layout, even if files are not flat.
It searches the repo recursively, stages the runtime files under the names the app expects,
builds the EXE with PyInstaller, then compiles the installer with Inno Setup.

How to use
----------
Put this builder folder either:
- inside the repo root, or
- beside the repo root and pass -RepoRoot

Then run:

    .\make_client_installer.ps1

Optional:

    .\make_client_installer.ps1 -RepoRoot "C:\Path\To\Windsor_Widget_Code"
    .\make_client_installer.ps1 -OdbcInstallerPath "C:\Installers\msodbcsql18.msi"
    .\make_client_installer.ps1 -InnoSetupCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

What it looks for
-----------------
Required:
- main_patched_status_yu.py or main.py
- ui_mainwindow.py
- shipments_window_ui.py / ui_shipments_window.py
- yu_order_workflow.py
- yu_order_review_export_test_window.py

Optional:
- month_year_picker.py
- requirements.txt
- any .ico file

Important
---------
If month_year_picker.py is not found and the app imports it, the EXE build may fail.
If requirements.txt is not found, the builder falls back to installing:
- PyInstaller
- PySide6
- pyodbc
- openpyxl

Installer safety
----------------
The installer forces a SQL connection test before leaving the SQL page,
and it will refuse to write client_config.json unless those exact tested
values are still the validated values.


Builder v2 fix:
- Corrected PyInstaller spec root resolution so staged_app is resolved under the builder folder, not C:\.
