from PySide6.QtCore import QDate, Signal
from PySide6.QtWidgets import QWidget, QHBoxLayout, QComboBox, QSizePolicy


MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

ABSOLUTE_YEAR_MIN = 2020
ABSOLUTE_YEAR_MAX = QDate.currentDate().year() + 5


class MonthYearPicker(QWidget):
    dateChanged = Signal(QDate)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.monthCombo = QComboBox(self)
        self.monthCombo.addItems(MONTH_NAMES)

        self.yearCombo = QComboBox(self)

        self._year_min = ABSOLUTE_YEAR_MIN
        self._year_max = ABSOLUTE_YEAR_MAX
        self._rebuild_years(self._year_min, self._year_max)

        control_height = 32

        # Fixed widths keep all picker rows aligned regardless of current text.
        self.monthCombo.setFixedWidth(140)
        self.yearCombo.setFixedWidth(110)

        self.monthCombo.setMinimumHeight(control_height)
        self.yearCombo.setMinimumHeight(control_height)

        self.monthCombo.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.yearCombo.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(self.monthCombo)
        layout.addWidget(self.yearCombo)
        layout.addStretch(1)

        today = QDate.currentDate()
        clamped_year = min(max(today.year(), self._year_min), self._year_max)
        today = QDate(clamped_year, today.month(), 1)
        self.setDate(today)

        self.monthCombo.currentIndexChanged.connect(self._emit_change)
        self.yearCombo.currentIndexChanged.connect(self._emit_change)

        combo_style = "QComboBox { padding-left: 6px; }"
        self.monthCombo.setStyleSheet(combo_style)
        self.yearCombo.setStyleSheet(combo_style)

    def _rebuild_years(self, minimum: int, maximum: int) -> None:
        minimum = max(ABSOLUTE_YEAR_MIN, minimum)
        maximum = min(ABSOLUTE_YEAR_MAX, maximum)
        if minimum > maximum:
            minimum, maximum = ABSOLUTE_YEAR_MIN, ABSOLUTE_YEAR_MAX

        current_year = None
        if hasattr(self, "yearCombo") and self.yearCombo.count():
            try:
                current_year = int(self.yearCombo.currentText())
            except ValueError:
                current_year = None

        self.yearCombo.blockSignals(True)
        self.yearCombo.clear()
        for year in range(minimum, maximum + 1):
            self.yearCombo.addItem(str(year))
        self.yearCombo.blockSignals(False)

        if current_year is not None and minimum <= current_year <= maximum:
            self.yearCombo.setCurrentText(str(current_year))
        else:
            self.yearCombo.setCurrentText(str(minimum))

        self._year_min = minimum
        self._year_max = maximum

    def date(self) -> QDate:
        return QDate(self.year(), self.month(), 1)

    def setDate(self, date: QDate) -> None:
        if not date.isValid():
            return

        if date.year() < self._year_min or date.year() > self._year_max:
            self._rebuild_years(min(date.year(), self._year_min), max(date.year(), self._year_max))

        year = min(max(date.year(), self._year_min), self._year_max)
        self.monthCombo.setCurrentIndex(date.month() - 1)
        self.yearCombo.setCurrentText(str(year))

    def year(self) -> int:
        return int(self.yearCombo.currentText())

    def month(self) -> int:
        return self.monthCombo.currentIndex() + 1

    def monthStart(self) -> QDate:
        return QDate(self.year(), self.month(), 1)

    def monthEnd(self) -> QDate:
        return QDate(self.year(), self.month(), 1).addMonths(1).addDays(-1)

    def setYearRange(self, minimum: int, maximum: int) -> None:
        self._rebuild_years(minimum, maximum)

    def _emit_change(self, *_):
        self.dateChanged.emit(self.date())
