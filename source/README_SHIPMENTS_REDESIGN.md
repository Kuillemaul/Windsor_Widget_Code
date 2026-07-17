# Windsor Widget v1.7.0

## Shipment Register Redesign

This release replaces the spreadsheet-style Shipments screen with a compact register and a dedicated detail panel.

### Permanent Shipment Numbers

Every shipment now receives a permanent Windsor shipment number:

```text
WT-SHP-000001
WT-SHP-000002
```

The shipment number is independent of:

- Purchase order number
- Forwarder shipment reference
- Container number
- Supplier

Existing shipment rows are assigned numbers automatically using their existing database record ID. New shipments receive a number when saved.

This gives the Widget a stable reference even when a PO changes or one shipment contains several purchase orders.

### New Shipment Screen

- Compact fixed-height shipment rows
- Muted status badges instead of full-row colours
- ETA warnings for overdue or near-due shipments
- Search by shipment number, supplier, PO, forwarder reference, container or vessel
- Quick views for:
  - Active
  - Needs Attention
  - Awaiting Booking
  - In Transit
  - Due Next 14 Days
  - YU Container
  - Arrived / Closed
  - Missing Information
- KPI cards showing the current shipment workload
- Right-side detail panel with dates, vessel, container, notes and update history

### Shipment Exceptions

The page identifies issues such as:

- Missing ETA
- Missing supplier
- Ready date passed without a booking
- ETD passed without being marked In Transit
- ETA passed without arrival confirmation
- Missing container number where expected
- Forwarder update older than seven days

### Separate Notes

Forwarder updates and Windsor internal notes are now stored separately.

Forwarder imports no longer overwrite internal notes.

### Change History

Shipment changes are logged with:

- Shipment number
- Date and time
- User
- Source of change
- Field changed
- Old and new values

### Import Improvements

- Import matching uses the permanent shipment number when it is present.
- Existing matching by forwarder reference, PO, supplier and container remains available.
- Import results show which fields changed.
- Newly created shipments show their assigned Windsor shipment numbers.
- Legacy `.xls` forwarder reports are supported through `xlrd`.

### Export Register

The current shipment view can be exported to CSV, including the permanent Windsor shipment number. This can be used for future exports, reconciliation and re-import matching.

### Database Changes

The app automatically adds the required shipment fields and creates:

```text
shipment_change_log
```

No manual SQL migration is required.

## Updated Files

```text
main_patched_status_yu.py
requirements.txt
```

## Installation

1. Close Windsor Widget.
2. Back up the Widget database and current source folder.
3. Replace `main_patched_status_yu.py`.
4. Update/install packages from `requirements.txt`.
5. Rebuild the application.
6. Open Shipments once to run the automatic database migration.
