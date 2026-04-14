# Windsor Widget Client Installer Build

This pack builds a **client installer** for the SQL Server version of Windsor Widget.

## What gets installed automatically
- The Windsor Widget app executable
- Bundled Python-side dependencies inside the EXE build:
  - PySide6
  - pyodbc
  - openpyxl
- `client_config.json` in `C:\ProgramData\WindsorWidget\client_config.json`
- The local customer files root in the current user's Qt settings / registry key
- The Windsor logo icon for the installer and app EXE

## What can also be installed automatically
If you stage the Microsoft **ODBC Driver 18 for SQL Server** installer into the `prereqs` folder before building, the client installer will install it when needed.

Supported staged filenames:
- `prereqs\msodbcsql18.msi`
- `prereqs\msodbcsql18.exe`

You can stage it automatically when building with:

```powershell
.\make_client_installer.ps1 -OdbcInstallerPath "C:\Installers\msodbcsql18.msi"
```

## Build steps
1. Install **Inno Setup 6** on the build PC.
2. Open PowerShell in this folder.
3. Run:

```powershell
.\make_client_installer.ps1
```

If Inno Setup is not found automatically:

```powershell
.\make_client_installer.ps1 -InnoSetupCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

## Clean rebuild

```powershell
.\make_client_installer.ps1 -Clean
```

## Output
The finished installer is created in:

```text
output\WindsorWidget_Client_1_0_0.exe
```

## Client install flow
During install you can:
- enter SQL Server settings
- **test the connection**
- choose the customer files folder

The installer writes only the SQL connection config globally. Customer file paths stay local per PC/user, while SQL stores only the matched **filename**.
