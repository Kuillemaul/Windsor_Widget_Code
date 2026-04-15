from __future__ import annotations

import csv
import sys
import importlib.util
import os
from pathlib import Path
from datetime import date, datetime

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QDoubleValidator
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCompleter,
    QDateEdit,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
)


class AddToOnOrderDialog(QDialog):
    def __init__(self, item_number: str, qty_text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add to On Order")
        self.resize(420, 220)

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel(f"Item: {item_number}", self))
        layout.addWidget(QLabel(f"Qty: {qty_text}", self))

        ready_row = QHBoxLayout()
        ready_row.addWidget(QLabel("Ready Date", self))
        self.ready_date_edit = QDateEdit(self)
        self.ready_date_edit.setCalendarPopup(True)
        self.ready_date_edit.setDisplayFormat("dd/MM/yy")
        self.ready_date_edit.setDate(QDate.currentDate())
        ready_row.addWidget(self.ready_date_edit, 1)
        layout.addLayout(ready_row)

        layout.addWidget(QLabel("Comments", self))
        self.comments_edit = QTextEdit(self)
        self.comments_edit.setPlaceholderText("Optional notes for YU / made-to-order / third-party delivery...")
        self.comments_edit.setMinimumHeight(90)
        layout.addWidget(self.comments_edit, 1)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        self.add_button = QPushButton("Add", self)
        self.add_button.clicked.connect(self.accept)
        button_row.addWidget(self.add_button)
        self.cancel_button = QPushButton("Cancel", self)
        self.cancel_button.clicked.connect(self.reject)
        button_row.addWidget(self.cancel_button)
        layout.addLayout(button_row)

    def values(self):
        return self.ready_date_edit.date().toString("dd/MM/yy"), self.comments_edit.toPlainText().strip()


class YUOrderEntryDialog(QDialog):
    def __init__(self, main_window, parent=None, initial_order_number="", initial_lines=None):
        super().__init__(parent or main_window)
        self.main_window = main_window
        self._loaded_draft_order_no = ""
        self._initial_lines = list(initial_lines or [])
        self.setWindowTitle("Create Order from YU")
        self.resize(860, 620)
        self.build_ui()
        self.setup_autocomplete()
        if initial_order_number:
            self.order_number_edit.setText(str(initial_order_number).strip())
            self._loaded_draft_order_no = str(initial_order_number).strip()
        if self._initial_lines:
            self.load_initial_lines(self._initial_lines)
        self.order_number_edit.setFocus()

    def build_ui(self):
        layout = QVBoxLayout(self)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Order Number", self))
        self.order_number_edit = QLineEdit(self)
        self.order_number_edit.setPlaceholderText("Enter order number")
        self.order_number_edit.editingFinished.connect(self.maybe_load_saved_draft)
        top_row.addWidget(self.order_number_edit, 1)
        layout.addLayout(top_row)

        entry_row = QHBoxLayout()
        entry_row.addWidget(QLabel("Item Number", self))
        self.item_number_edit = QLineEdit(self)
        self.item_number_edit.setPlaceholderText("Enter item number")
        entry_row.addWidget(self.item_number_edit, 2)

        entry_row.addWidget(QLabel("Qty", self))
        self.qty_edit = QLineEdit(self)
        self.qty_edit.setPlaceholderText("0")
        validator = QDoubleValidator(0.0, 999999999.0, 3, self.qty_edit)
        validator.setNotation(QDoubleValidator.StandardNotation)
        self.qty_edit.setValidator(validator)
        entry_row.addWidget(self.qty_edit, 1)

        self.add_line_button = QPushButton("Add Line", self)
        self.add_line_button.clicked.connect(self.add_current_line)
        entry_row.addWidget(self.add_line_button)
        layout.addLayout(entry_row)

        self.lines_table = QTableWidget(self)
        self.lines_table.setColumnCount(5)
        self.lines_table.setHorizontalHeaderLabels(["Item Number", "Description", "Qty", "Add To On Order", "Remove"])
        self.lines_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.lines_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.lines_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.horizontalHeader().setStretchLastSection(False)
        self.lines_table.cellDoubleClicked.connect(self.handle_table_double_click)
        layout.addWidget(self.lines_table, 1)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        self.save_later_button = QPushButton("Save for Later", self)
        self.save_later_button.clicked.connect(self.save_for_later)
        button_row.addWidget(self.save_later_button)

        self.validate_export_button = QPushButton("Validate and Export", self)
        self.validate_export_button.clicked.connect(self.validate_and_export)
        button_row.addWidget(self.validate_export_button)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked.connect(self.reject)
        button_row.addWidget(self.close_button)
        layout.addLayout(button_row)

    def load_initial_lines(self, lines):
        for line in lines or []:
            item_number = self.find_item_number(str(line.get("item_number", "")).strip())
            qty_value = float(line.get("qty", 0) or 0)
            if not item_number or qty_value <= 0:
                continue

            for row in range(self.lines_table.rowCount()):
                existing_item = self.lines_table.item(row, 0)
                if existing_item is not None and (existing_item.text() or "").strip() == item_number:
                    qty_item = self.lines_table.item(row, 2)
                    current_qty = float((qty_item.text() or '0').replace(',', '')) if qty_item is not None else 0.0
                    new_qty = current_qty + qty_value
                    self.lines_table.setItem(row, 2, QTableWidgetItem(self.format_qty(new_qty)))
                    return

            row = self.lines_table.rowCount()
            self.lines_table.insertRow(row)
            self.lines_table.setItem(row, 0, QTableWidgetItem(item_number))
            self.lines_table.setItem(row, 1, QTableWidgetItem(self.description_for_item(item_number)))
            self.lines_table.setItem(row, 2, QTableWidgetItem(self.format_qty(qty_value)))
            add_on_order_item = QTableWidgetItem('Add To On Order')
            add_on_order_item.setTextAlignment(Qt.AlignCenter)
            self.lines_table.setItem(row, 3, add_on_order_item)
            remove_item = QTableWidgetItem('Remove')
            remove_item.setTextAlignment(Qt.AlignCenter)
            self.lines_table.setItem(row, 4, remove_item)

        try:
            self.lines_table.resizeColumnsToContents()
        except Exception:
            pass

    def setup_autocomplete(self):
        item_numbers = getattr(self.main_window, 'item_numbers', []) or []
        completer = QCompleter(item_numbers, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchStartsWith)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.setModelSorting(QCompleter.CaseInsensitivelySortedModel)
        self.item_number_edit.setCompleter(completer)

    def normalised_order_number(self) -> str:
        return (self.order_number_edit.text() or '').strip()

    def draft_path_for_order(self, order_no: str) -> Path:
        return Path(self.main_window.get_yu_order_drafts_dir()) / f"{order_no}.csv"

    def maybe_load_saved_draft(self):
        order_no = self.normalised_order_number()
        if not order_no or order_no == self._loaded_draft_order_no or self.lines_table.rowCount() > 0:
            return
        draft_path = self.draft_path_for_order(order_no)
        if not draft_path.exists():
            return
        answer = QMessageBox.question(
            self,
            "Load saved draft",
            f"A saved YU draft exists for order {order_no}. Load it?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            return
        self.load_draft(draft_path)
        self._loaded_draft_order_no = order_no

    def find_item_number(self, typed_item: str) -> str:
        typed_item = (typed_item or '').strip()
        if not typed_item:
            return ''
        finder = getattr(self.main_window, 'find_item_number', None)
        if callable(finder):
            resolved = finder(typed_item)
            if resolved:
                return resolved
        item_numbers = getattr(self.main_window, 'item_numbers', []) or []
        lowered = typed_item.casefold()
        for item in item_numbers:
            if str(item).casefold() == lowered:
                return str(item)
        starts = [str(item) for item in item_numbers if str(item).casefold().startswith(lowered)]
        if len(starts) == 1:
            return starts[0]
        return typed_item

    def description_for_item(self, item_number: str) -> str:
        getter = getattr(self.main_window, 'get_item_master_row', None)
        get_first = getattr(self.main_window, 'get_first', None)
        if callable(getter) and callable(get_first):
            row = getter(item_number)
            return get_first(row, 'description', 'item_name', 'Item Name', 'Description') or ''
        return ''

    def add_current_line(self):
        order_no = self.normalised_order_number()
        if not order_no:
            QMessageBox.warning(self, 'Missing order number', 'Enter the order number first.')
            self.order_number_edit.setFocus()
            return

        item_number = self.find_item_number(self.item_number_edit.text())
        if not item_number:
            QMessageBox.warning(self, 'Missing item', 'Enter an item number.')
            self.item_number_edit.setFocus()
            return

        qty_text = (self.qty_edit.text() or '').strip()
        try:
            qty_value = float(qty_text)
        except Exception:
            qty_value = 0.0
        if qty_value <= 0:
            QMessageBox.warning(self, 'Invalid quantity', 'Enter a quantity greater than 0.')
            self.qty_edit.setFocus()
            self.qty_edit.selectAll()
            return

        for row in range(self.lines_table.rowCount()):
            existing_item = self.lines_table.item(row, 0)
            if existing_item is not None and (existing_item.text() or '').strip() == item_number:
                qty_item = self.lines_table.item(row, 2)
                current_qty = float((qty_item.text() or '0').replace(',', '')) if qty_item is not None else 0.0
                new_qty = current_qty + qty_value
                self.lines_table.setItem(row, 2, QTableWidgetItem(self.format_qty(new_qty)))
                self.item_number_edit.clear()
                self.qty_edit.clear()
                self.item_number_edit.setFocus()
                return

        row = self.lines_table.rowCount()
        self.lines_table.insertRow(row)
        self.lines_table.setItem(row, 0, QTableWidgetItem(item_number))
        self.lines_table.setItem(row, 1, QTableWidgetItem(self.description_for_item(item_number)))
        qty_item = QTableWidgetItem(self.format_qty(qty_value))
        qty_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lines_table.setItem(row, 2, qty_item)
        add_on_order_item = QTableWidgetItem('Add To On Order')
        add_on_order_item.setTextAlignment(Qt.AlignCenter)
        self.lines_table.setItem(row, 3, add_on_order_item)
        remove_item = QTableWidgetItem('Remove')
        remove_item.setTextAlignment(Qt.AlignCenter)
        self.lines_table.setItem(row, 4, remove_item)
        self.lines_table.resizeRowsToContents()
        self.lines_table.resizeColumnsToContents()
        self.item_number_edit.clear()
        self.qty_edit.clear()
        self.item_number_edit.setFocus()

    def handle_table_double_click(self, row: int, column: int):
        if row < 0:
            return
        if column == 3:
            self.add_row_to_on_order(row)
            return
        if column != 4:
            return
        item_number = self.lines_table.item(row, 0).text() if self.lines_table.item(row, 0) else 'this line'
        result = QMessageBox.question(
            self,
            'Remove line',
            f'Remove {item_number} from the YU order?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if result == QMessageBox.Yes:
            self.lines_table.removeRow(row)

    def add_row_to_on_order(self, row: int):
        if row < 0:
            return
        order_no = self.normalised_order_number()
        if not order_no:
            QMessageBox.warning(self, 'Missing order number', 'Enter the order number first.')
            self.order_number_edit.setFocus()
            return

        item_number = self.lines_table.item(row, 0).text().strip() if self.lines_table.item(row, 0) else ''
        qty_text = self.lines_table.item(row, 2).text().strip() if self.lines_table.item(row, 2) else '0'
        try:
            qty_value = float(qty_text.replace(',', ''))
        except Exception:
            qty_value = 0.0
        if not item_number or qty_value <= 0:
            QMessageBox.warning(self, 'Invalid line', 'This row does not contain a valid item and quantity.')
            return

        popup = AddToOnOrderDialog(item_number, qty_text, self)
        if popup.exec() != QDialog.Accepted:
            return
        ready_date_text, comments_text = popup.values()

        add_func = getattr(self.main_window, 'add_or_update_on_order_line', None)
        if not callable(add_func):
            QMessageBox.warning(self, 'On Order unavailable', 'The main window does not support adding to On Order.')
            return

        ok, message = add_func(order_no, item_number, qty_value, ready_date_text, comments_text)
        if ok:
            QMessageBox.information(self, 'Added to On Order', message)
        else:
            QMessageBox.warning(self, 'Add to On Order', message)

    def rows_as_dicts(self):
        rows = []
        order_no = self.normalised_order_number()
        today_text = date.today().strftime('%d/%m/%Y')
        for row in range(self.lines_table.rowCount()):
            item_number = self.lines_table.item(row, 0).text().strip() if self.lines_table.item(row, 0) else ''
            qty_text = self.lines_table.item(row, 2).text().strip() if self.lines_table.item(row, 2) else '0'
            try:
                qty_value = float(qty_text.replace(',', ''))
            except Exception:
                qty_value = 0.0
            if not item_number or qty_value <= 0:
                continue
            rows.append({
                'Date': today_text,
                'Order Number': order_no,
                'Item Number': item_number,
                'QTY': qty_value,
            })
        return rows

    def write_csv(self, path: Path):
        rows = self.rows_as_dicts()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open('w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Date', 'Order Number', 'Item Number', 'QTY'])
            writer.writeheader()
            writer.writerows(rows)

    def load_draft(self, draft_path: Path):
        self.lines_table.setRowCount(0)
        with draft_path.open(newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        if rows:
            self.order_number_edit.setText(str(rows[0].get('Order Number', '')).strip())
        for row_data in rows:
            row = self.lines_table.rowCount()
            self.lines_table.insertRow(row)
            self.lines_table.setItem(row, 0, QTableWidgetItem(str(row_data.get('Item Number', '')).strip()))
            self.lines_table.setItem(row, 1, QTableWidgetItem(self.description_for_item(str(row_data.get('Item Number', '')).strip())))
            qty_item = QTableWidgetItem(self.format_qty(float(row_data.get('QTY', 0) or 0)))
            qty_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.lines_table.setItem(row, 2, qty_item)
            add_on_order_item = QTableWidgetItem('Add To On Order')
            add_on_order_item.setTextAlignment(Qt.AlignCenter)
            self.lines_table.setItem(row, 3, add_on_order_item)
            remove_item = QTableWidgetItem('Remove')
            remove_item.setTextAlignment(Qt.AlignCenter)
            self.lines_table.setItem(row, 4, remove_item)
        self.lines_table.resizeRowsToContents()
        self.lines_table.resizeColumnsToContents()

    def save_for_later(self):
        order_no = self.normalised_order_number()
        rows = self.rows_as_dicts()
        if not order_no:
            QMessageBox.warning(self, 'Missing order number', 'Enter the order number first.')
            return
        if not rows:
            QMessageBox.warning(self, 'No lines', 'Add at least one line before saving.')
            return
        draft_path = self.draft_path_for_order(order_no)
        self.write_csv(draft_path)
        self._loaded_draft_order_no = order_no
        QMessageBox.information(self, 'Draft saved', f'Saved YU order draft to:\n{draft_path}')

    def validate_and_export(self):
        order_no = self.normalised_order_number()
        rows = self.rows_as_dicts()
        if not order_no:
            QMessageBox.warning(self, 'Missing order number', 'Enter the order number first.')
            return
        if not rows:
            QMessageBox.warning(self, 'No lines', 'Add at least one line before validating.')
            return
        temp_dir = Path(self.main_window.get_yu_order_temp_dir())
        temp_path = temp_dir / f'YU_{order_no}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        self.write_csv(temp_path)
        opened = self.main_window.open_yu_order_review_window(str(temp_path))
        if opened:
            self.accept()

    @staticmethod
    def format_qty(value: float) -> str:
        if float(value).is_integer():
            return str(int(value))
        return f'{value:,.3f}'.rstrip('0').rstrip('.')


def load_yu_review_module(main_window):
    module_name = 'yu_order_review_export_test_window'
    module_path = Path(__file__).resolve().parent / f'{module_name}.py'
    if not module_path.exists():
        raise FileNotFoundError(
            'yu_order_review_export_test_window.py was not found in the same folder as the app. '
            'Place that file beside the main app file first.'
        )

    existing = sys.modules.get(module_name)
    if existing is not None:
        return existing

    spec = importlib.util.spec_from_file_location(module_name, str(module_path))
    if spec is None or spec.loader is None:
        raise RuntimeError('Could not load yu_order_review_export_test_window.py')
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return module
