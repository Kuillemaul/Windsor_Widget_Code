YU Container PO Prototype

Purpose
-------
This prototype reads:
1. Packing List Worksheet.xlsx, using the PackingList sheet.
2. ITEMPURTEST.TXT, used as a mock MYOB AccountRight data source.

It does not write to MYOB.

Current logic
-------------
- Uses the Windsor item number already present in the packing list.
- Groups duplicate packing rows by Order No + Item Number.
- Looks up exact Order No + Item Number in the MYOB ITEMPUR export.
- Builds a grouped new PO preview using \ON header rows.
- Builds an adjustment preview for the original MYOB purchase orders.
- Flags oversupply where the packing list quantity is greater than the MYOB order line quantity.
- Flags empty/delete candidates where all lines on a touched MYOB PO would be reduced to zero.

Run
---
python yu_container_po_prototype.py

Recommended demo steps
----------------------
1. Browse to the Packing List Worksheet workbook.
2. Browse to ITEMPURTEST.TXT.
3. Click Run Prototype Analysis.
4. Review:
   - Output PO Preview
   - MYOB Adjustments
   - PO Summary
   - Issues
5. Use Export CSV Results if you want files for management review.

Important
---------
This is a prototype only. Later API implementation should:
- GET the current MYOB PO immediately before adjusting.
- Compare RowVersion/current quantities.
- Present a final approval screen.
- POST the new PO.
- PUT adjusted old POs.
- Never delete/close old POs automatically without approval.
