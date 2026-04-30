from __future__ import annotations

import csv
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from zipfile import ZipFile

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

DEFAULT_EXCHANGE_RATE = 0.68
DEFAULT_MAJOR_PRICE_THRESHOLD = 0.20


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        return str(int(float(text)))
    return text


def to_float(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except Exception:
        return 0.0


def fmt_qty(value):
    try:
        value = float(value)
    except Exception:
        return clean_text(value)
    if abs(value - round(value)) < 0.000001:
        return str(int(round(value)))
    return f"{value:.3f}".rstrip("0").rstrip(".")


def norm_order(value):
    text = clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    match = re.match(r"^0*(\d{5,6})( [A-Z]{1,4})?$", text)
    if match:
        return match.group(1) + (match.group(2) or "")
    return text


def norm_item(value):
    text = clean_text(value).upper()
    if text.startswith("\\"):
        return text.replace(" ", "")
    return re.sub(r"[^A-Z0-9]+", "", text)


def extract_order_from_description(description):
    text = clean_text(description).upper()
    match = re.search(r"O\s*/\s*NO\.?\s*([A-Z0-9 ]+)", text)
    if not match:
        return ""
    raw = match.group(1).strip()
    parts = raw.split()
    if not parts:
        return ""
    if len(parts) > 1 and re.match(r"^[A-Z]{1,4}$", parts[1]):
        return norm_order(parts[0] + " " + parts[1])
    return norm_order(parts[0])


def col_to_index(cell_ref):
    letters = "".join(ch for ch in str(cell_ref) if ch.isalpha())
    index = 0
    for ch in letters.upper():
        index = index * 26 + ord(ch) - 64
    return index - 1


def read_shared_strings(zip_file):
    try:
        root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values = []
    for si in root.findall("a:si", NS):
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(t.text or "")
        values.append("".join(parts))
    return values


def workbook_sheet_map(xlsx_path):
    with ZipFile(xlsx_path) as zip_file:
        workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
        relationships = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
        sheets = {}
        for sheet in workbook.find("a:sheets", NS):
            relationship_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_map.get(relationship_id, "")
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            sheets[sheet.attrib["name"]] = target
        return sheets


def read_xlsx_sheet(xlsx_path, sheet_name):
    sheet_map = workbook_sheet_map(xlsx_path)
    if sheet_name not in sheet_map:
        raise ValueError(f"Sheet '{sheet_name}' was not found. Available sheets: {', '.join(sheet_map)}")
    with ZipFile(xlsx_path) as zip_file:
        shared_strings = read_shared_strings(zip_file)
        root = ET.fromstring(zip_file.read(sheet_map[sheet_name]))

    rows = []
    max_col = 0
    for row in root.findall(".//a:sheetData/a:row", NS):
        row_index = int(row.attrib.get("r", "1")) - 1
        while len(rows) <= row_index:
            rows.append([])
        values = rows[row_index]
        for cell in row.findall("a:c", NS):
            col_index = col_to_index(cell.attrib.get("r", "A1"))
            while len(values) <= col_index:
                values.append(None)
            cell_type = cell.attrib.get("t")
            raw_value = cell.find("a:v", NS)
            inline_string = cell.find("a:is", NS)

            if cell_type == "s" and raw_value is not None:
                value = shared_strings[int(raw_value.text)]
            elif cell_type == "inlineStr" and inline_string is not None:
                value = "".join(t.text or "" for t in inline_string.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            elif raw_value is not None:
                value = raw_value.text
            else:
                value = None

            values[col_index] = value
            max_col = max(max_col, col_index + 1)

    for values in rows:
        values.extend([None] * (max_col - len(values)))
    return rows


def load_packing_list(path):
    rows = read_xlsx_sheet(path, "PackingList")
    items = []
    for row_number, row in enumerate(rows, start=1):
        if row_number < 24:
            continue

        supplier_item = clean_text(row[0] if len(row) > 0 else "")
        order_no_raw = clean_text(row[1] if len(row) > 1 else "")
        qty = to_float(row[7] if len(row) > 7 else 0)
        item_number = clean_text(row[12] if len(row) > 12 else "")

        if not supplier_item or not order_no_raw or not item_number or qty <= 0:
            continue

        items.append({
            "Source Row": row_number,
            "Order No": norm_order(order_no_raw),
            "Order No Display": order_no_raw,
            "Item Number": item_number,
            "Item Key": norm_item(item_number),
            "Supplier Item": supplier_item,
            "Size": clean_text(row[2] if len(row) > 2 else ""),
            "Colour": clean_text(row[3] if len(row) > 3 else ""),
            "Roll/Spool": clean_text(row[4] if len(row) > 4 else ""),
            "Mt/Unit": clean_text(row[5] if len(row) > 5 else ""),
            "Labelled As": clean_text(row[6] if len(row) > 6 else ""),
            "Qty": qty,
            "Packages": to_float(row[8] if len(row) > 8 else 0),
            "Net Weight": to_float(row[9] if len(row) > 9 else 0),
            "Gross Weight": to_float(row[10] if len(row) > 10 else 0),
            "CBM": clean_text(row[11] if len(row) > 11 else ""),
        })
    return items


def load_invoice_prices(path):
    try:
        rows = read_xlsx_sheet(path, "Invoice")
    except Exception:
        return {}

    grouped = {}
    for row_number, row in enumerate(rows, start=1):
        if row_number < 24:
            continue

        supplier_item = clean_text(row[0] if len(row) > 0 else "")
        order_no_raw = clean_text(row[1] if len(row) > 1 else "")
        unit_usd = to_float(row[7] if len(row) > 7 else 0)
        qty = to_float(row[8] if len(row) > 8 else 0)
        total_usd = to_float(row[9] if len(row) > 9 else 0)
        item_number = clean_text(row[10] if len(row) > 10 else "")

        if not supplier_item or not order_no_raw or not item_number or qty <= 0:
            continue

        if unit_usd <= 0 and total_usd > 0:
            unit_usd = total_usd / qty
        if total_usd <= 0 and unit_usd > 0:
            total_usd = unit_usd * qty

        key = (norm_order(order_no_raw), norm_item(item_number))
        if key not in grouped:
            grouped[key] = {
                "Order No": norm_order(order_no_raw),
                "Order No Display": order_no_raw,
                "Item Number": item_number,
                "Item Key": norm_item(item_number),
                "Supplier Item": supplier_item,
                "Invoice Qty": 0.0,
                "Invoice USD Total": 0.0,
                "Invoice Source Rows": [],
                "Invoice Unit Prices": [],
            }

        target = grouped[key]
        target["Invoice Qty"] += qty
        target["Invoice USD Total"] += total_usd
        target["Invoice Source Rows"].append(row_number)
        if unit_usd > 0:
            target["Invoice Unit Prices"].append(unit_usd)

    for target in grouped.values():
        if target["Invoice Qty"] > 0 and target["Invoice USD Total"] > 0:
            target["Invoice USD Unit"] = target["Invoice USD Total"] / target["Invoice Qty"]
        elif target["Invoice Unit Prices"]:
            target["Invoice USD Unit"] = sum(target["Invoice Unit Prices"]) / len(target["Invoice Unit Prices"])
        else:
            target["Invoice USD Unit"] = 0.0

        rounded_prices = {round(price, 6) for price in target["Invoice Unit Prices"]}
        target["Invoice Price Note"] = "Multiple invoice unit prices" if len(rounded_prices) > 1 else ""
        target["Invoice Source Rows"] = ",".join(str(row) for row in target["Invoice Source Rows"])

    return grouped


def format_price(value):
    try:
        return f"{float(value):.5f}".rstrip("0").rstrip(".")
    except Exception:
        return clean_text(value)


def format_percent(value):
    try:
        return f"{float(value) * 100:.1f}%"
    except Exception:
        return ""


def build_price_check(invoice_row, myob_price_aud, exchange_rate=DEFAULT_EXCHANGE_RATE, major_price_threshold=DEFAULT_MAJOR_PRICE_THRESHOLD):
    exchange_rate = float(exchange_rate or DEFAULT_EXCHANGE_RATE)
    major_price_threshold = float(major_price_threshold or DEFAULT_MAJOR_PRICE_THRESHOLD)
    myob_price_aud = float(myob_price_aud or 0.0)

    if not invoice_row:
        return {
            "Price Check": "NO INVOICE PRICE",
            "Invoice USD Unit": "",
            "Invoice AUD Unit": "",
            "MYOB AUD Unit": format_price(myob_price_aud),
            "Difference AUD": "",
            "Difference %": "",
            "Invoice Qty": "",
            "Invoice Source Rows": "",
            "Price Note": "No matching invoice row by order number + item number.",
            "_is_major_price_issue": False,
        }

    invoice_usd_unit = float(invoice_row.get("Invoice USD Unit") or 0.0)
    invoice_aud_unit = invoice_usd_unit / exchange_rate if exchange_rate else 0.0
    diff_aud = invoice_aud_unit - myob_price_aud
    diff_pct = diff_aud / myob_price_aud if myob_price_aud else 0.0

    if invoice_usd_unit <= 0 or myob_price_aud <= 0:
        status = "PRICE UNKNOWN"
        is_major = False
        note = "Missing invoice or MYOB unit price."
    elif abs(diff_pct) >= major_price_threshold:
        status = "PRICE ISSUE"
        is_major = True
        note = f"Major price difference at {format_percent(diff_pct)} using exchange rate {exchange_rate}."
    else:
        status = "PRICE OK"
        is_major = False
        note = invoice_row.get("Invoice Price Note") or "Within threshold."

    return {
        "Price Check": status,
        "Invoice USD Unit": format_price(invoice_usd_unit),
        "Invoice AUD Unit": format_price(invoice_aud_unit),
        "MYOB AUD Unit": format_price(myob_price_aud),
        "Difference AUD": format_price(diff_aud),
        "Difference %": format_percent(diff_pct),
        "Invoice Qty": fmt_qty(invoice_row.get("Invoice Qty", 0)),
        "Invoice Source Rows": clean_text(invoice_row.get("Invoice Source Rows")),
        "Price Note": note,
        "_is_major_price_issue": is_major,
    }


def load_myob_itempur(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as file:
        lines = file.readlines()

    header_index = 0
    for index, line in enumerate(lines):
        if "Purchase No." in line and "Item Number" in line:
            header_index = index
            break

    reader = csv.DictReader(lines[header_index:])
    rows = []
    current_header_order = ""
    current_header_purchase = ""

    for line_no, row in enumerate(reader, start=header_index + 2):
        if not any(clean_text(value) for value in row.values()):
            continue

        item_number = clean_text(row.get("Item Number"))
        description = clean_text(row.get("Description"))
        purchase_no = norm_order(row.get("Purchase No."))

        if purchase_no != current_header_purchase:
            current_header_order = ""
            current_header_purchase = purchase_no

        header_order = ""
        if norm_item(item_number) == "\\ON":
            current_header_order = extract_order_from_description(description)
            current_header_purchase = purchase_no
            header_order = current_header_order
        else:
            header_order = current_header_order

        rows.append({
            **row,
            "_line_no": line_no,
            "_purchase_no_norm": purchase_no,
            "_header_order_no": header_order,
            "_order_no": header_order or purchase_no,
            "_item_key": norm_item(item_number),
            "_is_header": norm_item(item_number) == "\\ON",
            "_qty": to_float(row.get("Quantity")),
            "_price": to_float(row.get("Price")),
        })

    return rows


def aggregate_packing(items):
    grouped = {}
    for row in items:
        key = (row["Order No"], row["Item Key"])
        if key not in grouped:
            grouped[key] = {
                "Order No": row["Order No"],
                "Order No Display": row["Order No Display"],
                "Item Number": row["Item Number"],
                "Item Key": row["Item Key"],
                "Supplier Item": row["Supplier Item"],
                "Size": row["Size"],
                "Colour": row["Colour"],
                "Labelled As": row["Labelled As"],
                "Packed Qty": 0.0,
                "Packages": 0.0,
                "Net Weight": 0.0,
                "Gross Weight": 0.0,
                "Source Rows": [],
            }
        target = grouped[key]
        target["Packed Qty"] += row["Qty"]
        target["Packages"] += row["Packages"]
        target["Net Weight"] += row["Net Weight"]
        target["Gross Weight"] += row["Gross Weight"]
        target["Source Rows"].append(row["Source Row"])
    return list(grouped.values())


def build_myob_indexes(myob_rows):
    by_order_item = defaultdict(list)
    by_po = defaultdict(list)
    for row in myob_rows:
        if row["_is_header"]:
            continue
        if not row["_item_key"] or row["_item_key"] in {"\\", "\\ON"}:
            continue
        by_order_item[(row["_order_no"], row["_item_key"])].append(row)
        by_po[row["_purchase_no_norm"]].append(row)
    return by_order_item, by_po


def analyse_container(packing_path, myob_path, exchange_rate=DEFAULT_EXCHANGE_RATE, major_price_threshold=DEFAULT_MAJOR_PRICE_THRESHOLD):
    packing_items = load_packing_list(packing_path)
    grouped_packing = aggregate_packing(packing_items)
    invoice_prices = load_invoice_prices(packing_path)
    myob_rows = load_myob_itempur(myob_path)
    by_order_item, by_po = build_myob_indexes(myob_rows)

    adjustments = []
    issues = []
    price_checks = []
    po_rows = []
    orders_touched = set()
    oversupply_lines = 0
    zeroed_lines = 0
    duplicate_review_lines = 0

    for packed in grouped_packing:
        candidates = by_order_item.get((packed["Order No"], packed["Item Key"]), [])
        chosen = None
        status = "MATCHED"
        note = ""

        if len(candidates) == 1:
            chosen = candidates[0]
        elif len(candidates) > 1:
            sufficient = [row for row in candidates if row["_qty"] >= packed["Packed Qty"]]
            if len(sufficient) == 1:
                chosen = sufficient[0]
                status = "MATCHED_MULTIPLE_RESOLVED"
                note = "Multiple MYOB lines; selected the only line with enough quantity."
            elif sufficient:
                chosen = sorted(sufficient, key=lambda row: row["_qty"])[0]
                status = "REVIEW_DUPLICATE"
                note = "Multiple candidate MYOB lines. Prototype picked the smallest sufficient line."
                duplicate_review_lines += 1
            else:
                chosen = sorted(candidates, key=lambda row: row["_qty"], reverse=True)[0]
                status = "OVERSUPPLY"
                note = "Multiple candidate lines, but all are under the packed quantity."
                duplicate_review_lines += 1

        if chosen is None:
            issues.append({
                "Issue Type": "UNMATCHED",
                "Order No": packed["Order No Display"],
                "Item Number": packed["Item Number"],
                "Packed Qty": packed["Packed Qty"],
                "MYOB Qty": "",
                "Difference": "",
                "Supplier Item": packed["Supplier Item"],
                "Colour": packed["Colour"],
                "Notes": "No MYOB line found for order number + item number.",
            })
            continue

        original_qty = chosen["_qty"]
        packed_qty = packed["Packed Qty"]
        remaining_qty = original_qty - packed_qty

        if remaining_qty < -0.000001:
            status = "OVERSUPPLY"
            oversupply_lines += 1
            issues.append({
                "Issue Type": "OVERSUPPLY",
                "Order No": packed["Order No Display"],
                "Item Number": packed["Item Number"],
                "Packed Qty": packed_qty,
                "MYOB Qty": original_qty,
                "Difference": packed_qty - original_qty,
                "Supplier Item": packed["Supplier Item"],
                "Colour": packed["Colour"],
                "Notes": "Packed quantity is higher than MYOB quantity. Do not auto-adjust below zero.",
            })
        elif abs(remaining_qty) <= 0.000001:
            zeroed_lines += 1

        orders_touched.add(chosen["_purchase_no_norm"] or packed["Order No"])

        invoice_row = invoice_prices.get((packed["Order No"], packed["Item Key"]))
        price_check = build_price_check(invoice_row, chosen["_price"], exchange_rate, major_price_threshold)
        if price_check["_is_major_price_issue"]:
            issues.append({
                "Issue Type": "PRICE DISCREPANCY",
                "Order No": packed["Order No Display"],
                "Item Number": packed["Item Number"],
                "Packed Qty": packed_qty,
                "MYOB Qty": original_qty,
                "Difference": price_check["Difference %"],
                "Supplier Item": packed["Supplier Item"],
                "Colour": packed["Colour"],
                "Notes": f"Invoice USD unit {price_check['Invoice USD Unit']} converts to AUD {price_check['Invoice AUD Unit']} at {exchange_rate}; MYOB AUD unit is {price_check['MYOB AUD Unit']}.",
            })

        price_checks.append({
            "Price Check": price_check["Price Check"],
            "Supplier": clean_text(chosen.get("Co./Last Name")),
            "MYOB Purchase No": clean_text(chosen.get("Purchase No.")),
            "Order No": packed["Order No Display"],
            "Item Number": clean_text(chosen.get("Item Number")),
            "Description": clean_text(chosen.get("Description")),
            "Invoice USD Unit": price_check["Invoice USD Unit"],
            "Invoice AUD Unit": price_check["Invoice AUD Unit"],
            "MYOB AUD Unit": price_check["MYOB AUD Unit"],
            "Difference AUD": price_check["Difference AUD"],
            "Difference %": price_check["Difference %"],
            "Invoice Qty": price_check["Invoice Qty"],
            "Packed Qty": packed_qty,
            "Original MYOB Qty": original_qty,
            "Invoice Source Rows": price_check["Invoice Source Rows"],
            "Packing Source Rows": ",".join(str(row) for row in packed["Source Rows"]),
            "Note": price_check["Price Note"],
        })

        adjustments.append({
            "Status": status,
            "Supplier": clean_text(chosen.get("Co./Last Name")),
            "MYOB Purchase No": clean_text(chosen.get("Purchase No.")),
            "Order No": packed["Order No Display"],
            "Item Number": clean_text(chosen.get("Item Number")),
            "Description": clean_text(chosen.get("Description")),
            "Original MYOB Qty": original_qty,
            "Packed Qty": packed_qty,
            "Remaining Qty": max(0.0, remaining_qty),
            "Adjustment Qty": min(original_qty, packed_qty),
            "MYOB AUD Unit": price_check["MYOB AUD Unit"],
            "Invoice USD Unit": price_check["Invoice USD Unit"],
            "Invoice AUD Unit": price_check["Invoice AUD Unit"],
            "Price Difference %": price_check["Difference %"],
            "Price Check": price_check["Price Check"],
            "Source Rows": ",".join(str(row) for row in packed["Source Rows"]),
            "Line No": chosen["_line_no"],
            "Note": note,
        })

        po_rows.append({
            "Order No": packed["Order No Display"],
            "Item Number": packed["Item Number"],
            "Quantity": packed_qty,
            "Description": clean_text(chosen.get("Description")),
            "Price": clean_text(chosen.get("Price")),
            "Invoice USD Unit": price_check["Invoice USD Unit"],
            "Invoice AUD Unit": price_check["Invoice AUD Unit"],
            "Price Check": price_check["Price Check"],
            "Supplier Item": packed["Supplier Item"],
            "Colour": packed["Colour"],
            "Source Rows": ",".join(str(row) for row in packed["Source Rows"]),
            "Status": status,
        })

    adjust_by_line = defaultdict(float)
    for adjustment in adjustments:
        adjust_by_line[int(adjustment["Line No"])] += float(adjustment["Adjustment Qty"])

    po_summary = []
    for po_key, lines in by_po.items():
        if po_key not in orders_touched:
            continue
        before_total = 0.0
        after_total = 0.0
        adjusted_lines = 0
        zeroed_po_lines = 0
        for line in lines:
            before = line["_qty"]
            after = max(0.0, before - adjust_by_line.get(line["_line_no"], 0.0))
            before_total += before
            after_total += after
            if abs(after - before) > 0.000001:
                adjusted_lines += 1
            if before > 0 and after <= 0.000001:
                zeroed_po_lines += 1

        first = lines[0]
        po_summary.append({
            "MYOB Purchase No": clean_text(first.get("Purchase No.")),
            "Supplier": clean_text(first.get("Co./Last Name")),
            "Lines": len(lines),
            "Adjusted Lines": adjusted_lines,
            "Zeroed Lines": zeroed_po_lines,
            "Before Qty": before_total,
            "After Qty": after_total,
            "Qty Reduced": before_total - after_total,
            "Empty/Delete Candidate": "YES" if after_total <= 0.000001 else "NO",
        })

    po_rows.sort(key=lambda row: (norm_order(row["Order No"]), norm_item(row["Item Number"])))
    adjustments.sort(key=lambda row: (norm_order(row["Order No"]), norm_item(row["Item Number"])))
    price_checks.sort(key=lambda row: (row["Price Check"] != "PRICE ISSUE", norm_order(row["Order No"]), norm_item(row["Item Number"])))
    po_summary.sort(key=lambda row: clean_text(row["MYOB Purchase No"]))

    po_preview = []
    last_order = None
    line_no = 0
    for row in po_rows:
        if row["Order No"] != last_order:
            line_no += 1
            po_preview.append({
                "Line": line_no,
                "Line Type": "HEADER",
                "Order No": row["Order No"],
                "Item Number": "\\ON",
                "Quantity": 0,
                "Description": f"O/NO {row['Order No']}",
                "Price": "$0.00",
                "Invoice USD Unit": "",
                "Invoice AUD Unit": "",
                "Price Check": "",
                "Status": "GROUP HEADER",
            })
            last_order = row["Order No"]
        line_no += 1
        po_preview.append({
            "Line": line_no,
            "Line Type": "ITEM",
            "Order No": row["Order No"],
            "Item Number": row["Item Number"],
            "Quantity": row["Quantity"],
            "Description": row["Description"],
            "Price": row["Price"],
            "Invoice USD Unit": row.get("Invoice USD Unit", ""),
            "Invoice AUD Unit": row.get("Invoice AUD Unit", ""),
            "Price Check": row.get("Price Check", ""),
            "Status": row["Status"],
        })

    stats = {
        "Packing source rows": len(packing_items),
        "Grouped packing lines": len(grouped_packing),
        "New PO item lines": len(po_rows),
        "New PO visible lines incl. headers": len(po_preview),
        "MYOB POs adjusted": len(orders_touched),
        "MYOB lines adjusted": len(adjustments),
        "Lines reduced to zero": zeroed_lines,
        "Empty / delete candidate POs": sum(1 for po in po_summary if po["Empty/Delete Candidate"] == "YES"),
        "Oversupply lines": oversupply_lines,
        "Unmatched lines": sum(1 for issue in issues if issue["Issue Type"] == "UNMATCHED"),
        "Duplicate review lines": duplicate_review_lines,
        "Invoice price rows": len(invoice_prices),
        "Major price discrepancies": sum(1 for row in price_checks if row["Price Check"] == "PRICE ISSUE"),
        "Total packed qty": sum(row["Quantity"] for row in po_rows),
        "Total adjusted qty": sum(row["Adjustment Qty"] for row in adjustments),
        "Remaining qty on touched POs": sum(row["After Qty"] for row in po_summary),
    }

    packing_table = []
    for row in packing_items:
        packing_table.append({
            "Source Row": row["Source Row"],
            "Order No": row["Order No Display"],
            "Item Number": row["Item Number"],
            "Supplier Item": row["Supplier Item"],
            "Size": row["Size"],
            "Colour": row["Colour"],
            "Labelled As": row["Labelled As"],
            "Qty": row["Qty"],
            "Packages": row["Packages"],
            "Net Weight": row["Net Weight"],
            "Gross Weight": row["Gross Weight"],
        })

    return {
        "stats": stats,
        "packing_lines": packing_table,
        "po_preview": po_preview,
        "adjustments": adjustments,
        "po_summary": po_summary,
        "issues": issues,
        "price_checks": price_checks,
    }


def write_csv(path, rows, fieldnames=None):
    fieldnames = fieldnames or (list(rows[0].keys()) if rows else [])
    with open(path, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


class StatCard(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setObjectName("statCard")
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(
            "QFrame#statCard {"
            " border: 1px solid #3B82F6;"
            " border-radius: 8px;"
            " background: #172033;"
            "}"
        )
        layout = QVBoxLayout(self)
        self.title = QLabel(title)
        self.title.setStyleSheet("font-size: 10pt; color: #BFDBFE; background: transparent; border: none;")
        self.value = QLabel("—")
        self.value.setStyleSheet("font-size: 15pt; font-weight: bold; color: #FFFFFF; background: transparent; border: none;")
        layout.addWidget(self.title)
        layout.addWidget(self.value)

    def set_value(self, value):
        self.value.setText(fmt_qty(value))


class PrototypeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YU Container Prototype - MYOB PO Preview")
        self.resize(1500, 900)
        self.results = None
        self.apply_readable_dark_theme()
        self.build_ui()

    def apply_readable_dark_theme(self):
        # Keep this prototype readable even when Windows / Qt is in dark mode.
        # The old pale warning colours were being paired with white text by the OS theme.
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #111827;
                color: #F9FAFB;
                font-size: 10pt;
            }
            QLabel {
                color: #F9FAFB;
                background: transparent;
            }
            QLineEdit, QTextEdit {
                background-color: #1F2937;
                color: #F9FAFB;
                border: 1px solid #4B5563;
                border-radius: 4px;
                padding: 4px;
                selection-background-color: #2563EB;
            }
            QPushButton {
                background-color: #2563EB;
                color: #FFFFFF;
                border: 1px solid #60A5FA;
                border-radius: 5px;
                padding: 6px 10px;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            QPushButton:pressed {
                background-color: #1E40AF;
            }
            QTabWidget::pane {
                border: 1px solid #4B5563;
                background-color: #111827;
            }
            QTabBar::tab {
                background-color: #1F2937;
                color: #D1D5DB;
                border: 1px solid #374151;
                padding: 7px 12px;
            }
            QTabBar::tab:selected {
                background-color: #2563EB;
                color: #FFFFFF;
                border-color: #60A5FA;
            }
            QTableWidget {
                background-color: #111827;
                alternate-background-color: #1F2937;
                color: #F9FAFB;
                gridline-color: #374151;
                border: 1px solid #4B5563;
                selection-background-color: #2563EB;
                selection-color: #FFFFFF;
            }
            QHeaderView::section {
                background-color: #374151;
                color: #FFFFFF;
                border: 1px solid #4B5563;
                padding: 5px;
                font-weight: bold;
            }
            QScrollBar:horizontal, QScrollBar:vertical {
                background: #111827;
                border: 1px solid #374151;
            }
            QScrollBar::handle:horizontal, QScrollBar::handle:vertical {
                background: #6B7280;
                border-radius: 4px;
                min-height: 24px;
                min-width: 24px;
            }
            QScrollBar::handle:horizontal:hover, QScrollBar::handle:vertical:hover {
                background: #9CA3AF;
            }
        """)

    def build_ui(self):
        root = QWidget()
        root_layout = QVBoxLayout(root)

        title = QLabel("YU Container Prototype")
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        root_layout.addWidget(title)

        info = QLabel("This prototype uses the PackingList sheet with Windsor item numbers already included. ITEMPURTEST.TXT is used as the mock MYOB data source. Nothing writes back to MYOB.")
        info.setWordWrap(True)
        root_layout.addWidget(info)

        picker_layout = QGridLayout()
        self.packing_edit = QLineEdit()
        self.myob_edit = QLineEdit()
        self.new_po_edit = QLineEdit("NEW_CONTAINER_PO")
        self.exchange_rate_edit = QLineEdit(str(DEFAULT_EXCHANGE_RATE))
        self.price_threshold_edit = QLineEdit(str(int(DEFAULT_MAJOR_PRICE_THRESHOLD * 100)))

        browse_packing = QPushButton("Browse Packing List")
        browse_packing.clicked.connect(self.browse_packing)
        browse_myob = QPushButton("Browse MYOB ITEMPUR")
        browse_myob.clicked.connect(self.browse_myob)

        run_button = QPushButton("Run Prototype Analysis")
        run_button.clicked.connect(self.run_analysis)
        run_button.setStyleSheet("font-weight: bold; padding: 6px 10px; background-color: #16A34A; border-color: #86EFAC;")

        export_button = QPushButton("Export CSV Results")
        export_button.clicked.connect(self.export_results)

        picker_layout.addWidget(QLabel("Packing List Workbook"), 0, 0)
        picker_layout.addWidget(self.packing_edit, 0, 1)
        picker_layout.addWidget(browse_packing, 0, 2)
        picker_layout.addWidget(QLabel("MYOB Data / Mock API GET"), 1, 0)
        picker_layout.addWidget(self.myob_edit, 1, 1)
        picker_layout.addWidget(browse_myob, 1, 2)
        picker_layout.addWidget(QLabel("New Container PO No"), 2, 0)
        picker_layout.addWidget(self.new_po_edit, 2, 1)
        picker_layout.addWidget(QLabel("USD→AUD factor"), 3, 0)
        picker_layout.addWidget(self.exchange_rate_edit, 3, 1)
        picker_layout.addWidget(QLabel("Major price threshold %"), 3, 2)
        picker_layout.addWidget(self.price_threshold_edit, 3, 3)
        picker_layout.addWidget(run_button, 4, 2)
        picker_layout.addWidget(export_button, 4, 3)
        root_layout.addLayout(picker_layout)

        self.stat_grid = QGridLayout()
        stat_names = [
            "Packing source rows",
            "Grouped packing lines",
            "New PO item lines",
            "MYOB POs adjusted",
            "MYOB lines adjusted",
            "Lines reduced to zero",
            "Empty / delete candidate POs",
            "Oversupply lines",
            "Unmatched lines",
            "Major price discrepancies",
            "Invoice price rows",
            "Total packed qty",
            "Total adjusted qty",
            "Remaining qty on touched POs",
        ]
        self.stat_cards = {}
        for index, name in enumerate(stat_names):
            card = StatCard(name)
            self.stat_cards[name] = card
            self.stat_grid.addWidget(card, index // 6, index % 6)
        root_layout.addLayout(self.stat_grid)

        self.tabs = QTabWidget()
        self.tables = {}
        for name in ["Output PO Preview", "MYOB Adjustments", "Price Check", "PO Summary", "Issues", "Packing Lines"]:
            table = QTableWidget()
            table.setAlternatingRowColors(True)
            table.setSortingEnabled(True)
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            self.tables[name] = table
            self.tabs.addTab(table, name)
        root_layout.addWidget(self.tabs, 1)

        notes = QTextEdit()
        notes.setReadOnly(True)
        notes.setMaximumHeight(90)
        notes.setPlainText("Process: Import PackingList + Invoice → read item/order/packed qty and invoice USD unit price → convert USD to AUD using the factor above → compare against MYOB order AUD price → preview new grouped PO and old PO reductions. Oversupply and major price issues are shown as review issues and should not be live-written without approval.")
        root_layout.addWidget(notes)

        self.setCentralWidget(root)

        base_dir = Path(__file__).resolve().parent
        possible_packing = base_dir / "Packing List Worksheet(1).xlsx"
        possible_myob = base_dir / "ITEMPURTEST(1).TXT"
        if possible_packing.exists():
            self.packing_edit.setText(str(possible_packing))
        if possible_myob.exists():
            self.myob_edit.setText(str(possible_myob))

    def browse_packing(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select packing list workbook", "", "Excel Workbooks (*.xlsx)")
        if path:
            self.packing_edit.setText(path)

    def browse_myob(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select MYOB ITEMPUR export", "", "MYOB Text Files (*.txt *.csv);;All Files (*)")
        if path:
            self.myob_edit.setText(path)

    def run_analysis(self):
        packing_path = self.packing_edit.text().strip()
        myob_path = self.myob_edit.text().strip()

        if not packing_path or not Path(packing_path).exists():
            QMessageBox.warning(self, "Missing packing list", "Select a valid Packing List Worksheet workbook.")
            return
        if not myob_path or not Path(myob_path).exists():
            QMessageBox.warning(self, "Missing MYOB data", "Select a valid ITEMPURTEST.TXT file.")
            return

        try:
            exchange_rate = to_float(self.exchange_rate_edit.text()) or DEFAULT_EXCHANGE_RATE
            threshold_value = to_float(self.price_threshold_edit.text())
            major_price_threshold = threshold_value / 100.0 if threshold_value > 1 else (threshold_value or DEFAULT_MAJOR_PRICE_THRESHOLD)
            self.results = analyse_container(packing_path, myob_path, exchange_rate=exchange_rate, major_price_threshold=major_price_threshold)
        except Exception as exc:
            QMessageBox.critical(self, "Analysis failed", str(exc))
            return

        for name, card in self.stat_cards.items():
            card.set_value(self.results["stats"].get(name, "—"))

        self.populate_table("Output PO Preview", self.results["po_preview"])
        self.populate_table("MYOB Adjustments", self.results["adjustments"])
        self.populate_table("Price Check", self.results["price_checks"])
        self.populate_table("PO Summary", self.results["po_summary"])
        self.populate_table("Issues", self.results["issues"])
        self.populate_table("Packing Lines", self.results["packing_lines"])

    def populate_table(self, name, rows):
        table = self.tables[name]
        table.setSortingEnabled(False)
        table.clear()

        if not rows:
            table.setRowCount(0)
            table.setColumnCount(0)
            table.setSortingEnabled(True)
            return

        columns = list(rows[0].keys())
        table.setColumnCount(len(columns))
        table.setRowCount(len(rows))
        table.setHorizontalHeaderLabels(columns)

        for row_index, row in enumerate(rows):
            row_status = clean_text(row.get("Status") or row.get("Issue Type") or row.get("Empty/Delete Candidate"))
            price_status = clean_text(row.get("Price Check"))
            for col_index, column in enumerate(columns):
                value = row.get(column, "")
                if isinstance(value, float):
                    display = fmt_qty(value)
                else:
                    display = clean_text(value)
                item = QTableWidgetItem(display)
                item.setFlags((item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled) & ~Qt.ItemIsEditable)
                # Use dark warning colours with explicit foregrounds.
                # This avoids unreadable white-on-pale-pink rows under Windows dark mode.
                if "OVERSUPPLY" in row_status:
                    item.setBackground(QBrush(QColor("#7F1D1D")))
                    item.setForeground(QBrush(QColor("#FFFFFF")))
                elif price_status == "PRICE ISSUE" or "PRICE DISCREPANCY" in row_status:
                    item.setBackground(QBrush(QColor("#831843")))
                    item.setForeground(QBrush(QColor("#FFFFFF")))
                elif "UNMATCHED" in row_status:
                    item.setBackground(QBrush(QColor("#7C2D12")))
                    item.setForeground(QBrush(QColor("#FFFFFF")))
                elif row_status == "YES":
                    item.setBackground(QBrush(QColor("#713F12")))
                    item.setForeground(QBrush(QColor("#FFFFFF")))
                elif row_status == "GROUP HEADER":
                    item.setBackground(QBrush(QColor("#1D4ED8")))
                    item.setForeground(QBrush(QColor("#FFFFFF")))
                else:
                    item.setForeground(QBrush(QColor("#F9FAFB")))
                table.setItem(row_index, col_index, item)

        table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    def export_results(self):
        if not self.results:
            QMessageBox.information(self, "No results", "Run the prototype analysis first.")
            return

        folder = QFileDialog.getExistingDirectory(self, "Select export folder")
        if not folder:
            return

        folder_path = Path(folder)
        write_csv(folder_path / "output_po_preview.csv", self.results["po_preview"])
        write_csv(folder_path / "myob_order_adjustments.csv", self.results["adjustments"])
        write_csv(folder_path / "price_check.csv", self.results["price_checks"])
        write_csv(folder_path / "myob_po_summary.csv", self.results["po_summary"])
        write_csv(folder_path / "issues_oversupply_unmatched_price.csv", self.results["issues"], ["Issue Type", "Order No", "Item Number", "Packed Qty", "MYOB Qty", "Difference", "Supplier Item", "Colour", "Notes"])
        write_csv(folder_path / "packing_lines_loaded.csv", self.results["packing_lines"])
        with open(folder_path / "stats.json", "w", encoding="utf-8") as file:
            json.dump(self.results["stats"], file, indent=2)

        QMessageBox.information(self, "Export complete", f"CSV results exported to:\n{folder_path}")


def main():
    app = QApplication(sys.argv)
    window = PrototypeWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
