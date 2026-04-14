# Windsor Widget Code

Windsor Widget is a Windows desktop application built for internal purchasing, stock, and container-planning workflows.

It brings several day-to-day tasks into one place, including customer and item summaries, order planning, container building, shipment tracking, customer file review, and YU supplier order export/review.

This repository contains the current working code and UI files for the app.

---

## What this app does

Windsor Widget is designed to support operational and purchasing workflows such as:

- reviewing customer purchase history
- reviewing item-level sales and stock position
- planning what needs to be ordered
- building container sheets
- tracking incoming shipments
- reviewing SABA-related purchasing/customer activity
- importing updated sales, stock, and order data
- reviewing and exporting YU supplier order forms from CSV input

This is not a general-purpose package. It is a business-specific desktop tool.

---

## Main features

### Customer Summary
- search and load customer data
- filter by month/year ranges
- preview linked customer files
- open customer files in Excel
- combine related customer/state accounts where needed

### Item Summary
- load item-level history and stock context
- show total quantity sold, average monthly sales, suggested minimums, stock on hand, stock on order, and related planning fields
- edit selected item master values directly from the UI by double-clicking editable fields such as:
  - Roll / Spool
  - Mt / Unit
  - Box
  - Pallet / Carton

### To Order Sheet
- enter items and quantities for purchasing review
- use item autocompletion
- support supplier selection and urgency flags
- feed planning workflows from the current item and stock context

### Build Container Sheet
- create and manage container lines
- add quantities, carton counts, notes, and flags
- export container data to Excel
- create an Outlook draft email with the export attached

### Shipments Window
- manage shipment rows for:
  - Melbourne
  - Sydney
  - SABA
- filter by shipment type
- edit shipment details directly in the table
- colour-code rows by shipment status and overdue dates
- open MarineTraffic or findTEU from shipment rows

### SABA Review
- review SABA-related customer/product activity from the main interface

### Update Data
- import and refresh:
  - sales
  - stock
  - orders

### YU Order Review / Export
- generate a YU order input CSV from the app
- launch a separate review workflow before export
- audit item mapping against the matched YU workbook
- manually override unresolved mappings
- export final supplier workbook outputs once review is complete

---

## Tech stack

- **Python**
- **PySide6**
- **openpyxl**
- **pyodbc**
- **SQL Server** for the current runtime build

---

## Current architecture

The current build is **SQL Server only**.

The app loads its database settings from `client_config.json` and expects the provider to be set to `sqlserver`. It can also read config from an override path via the `WINDSOR_WIDGET_CONFIG` environment variable.

Typical config search locations include:

- app folder
- `data/` under the app folder
- `C:\ProgramData\WindsorWidget\client_config.json`
- path supplied by `WINDSOR_WIDGET_CONFIG`

> Note: do **not** commit real usernames, passwords, or production server details to GitHub.

Example `client_config.json`:

```json
{
  "provider": "sqlserver",
  "driver": "ODBC Driver 18 for SQL Server",
  "server": "YOUR_SERVER",
  "port": "14330",
  "database": "WindsorWidget",
  "username": "YOUR_USERNAME",
  "password": "YOUR_PASSWORD",
  "trusted_connection": false,
  "encrypt": false,
  "trust_server_certificate": true,
  "timeout": 5
}
