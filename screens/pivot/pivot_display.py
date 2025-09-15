import pandas as pd
from PyQt6.QtWidgets import QTableView, QVBoxLayout, QFrame, QLabel, QHBoxLayout, QPushButton, QLineEdit, QComboBox
from PyQt6.QtCore import QSortFilterProxyModel, Qt, QAbstractTableModel, QModelIndex, pyqtSignal
from PyQt6.QtGui import QFont, QPalette, QColor
import logging

logger = logging.getLogger(__name__)

class PivotTableModel(QAbstractTableModel):
    dataChanged = pyqtSignal(QModelIndex, QModelIndex)

    def __init__(self, df=None):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()
        self._column_widths = {}

    def set_data(self, df):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return self._df.shape[0]

    def columnCount(self, parent=QModelIndex()):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not 0 <= index.row() < self._df.shape[0] or not 0 <= index.column() < self._df.shape[1]:
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._df.iloc[index.row(), index.column()]
            return str(value) if pd.notna(value) else ""
        elif role == Qt.ItemDataRole.BackgroundRole:
            # Custom background for rows (odd/even)
            if index.row() % 2 == 0:
                return QColor("#f9f9f9")
            return QColor("white")
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            col = index.column()
            if col == 0:  # Solution Label
                return Qt.AlignmentFlag.AlignLeft.value()
            return Qt.AlignmentFlag.AlignCenter.value()
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._df.columns[section])
            elif orientation == Qt.Orientation.Vertical:
                return str(section + 1)
        return None

    def set_column_width(self, col, width):
        self._column_widths[col] = width

class PivotDisplayWidget(QFrame):
    def __init__(self, parent, pivot_tab):
        super().__init__(parent)
        self.pivot_tab = pivot_tab
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Control frame
        control_frame = QFrame()
        control_layout = QHBoxLayout(control_frame)
        control_layout.setContentsMargins(5, 5, 5, 5)

        # Decimal places
        QLabel("Decimal Places:", control_frame)
        self.decimal_places = QComboBox()
        self.decimal_places.addItems(["0", "1", "2", "3"])
        self.decimal_places.setCurrentText("1")
        self.decimal_places.currentTextChanged.connect(self.pivot_tab.update_pivot_display)
        control_layout.addWidget(self.decimal_places)

        # Use Int checkbox
        self.use_int_var = QCheckBox("Use Int")
        self.use_int_var.toggled.connect(lambda: self.pivot_tab.create_pivot())
        control_layout.addWidget(self.use_int_var)

        # Diff range
        QLabel("Diff Range (%):", control_frame)
        self.diff_min = QLineEdit("-12")
        self.diff_min.textChanged.connect(self.pivot_tab.validate_diff_range)
        control_layout.addWidget(self.diff_min)
        QLabel("to", control_frame)
        self.diff_max = QLineEdit("12")
        self.diff_max.textChanged.connect(self.pivot_tab.validate_diff_range)
        control_layout.addWidget(self.diff_max)

        # Search
        self.search_var = QLineEdit()
        self.search_var.setPlaceholderText("Search...")
        self.search_var.textChanged.connect(self.pivot_tab.update_pivot_display)
        control_layout.addWidget(self.search_var)

        # Row filter button
        row_filter_btn = QPushButton("Row Filter")
        row_filter_btn.clicked.connect(self.pivot_tab.open_row_filter_window)
        control_layout.addWidget(row_filter_btn)

        # Column filter button
        col_filter_btn = QPushButton("Column Filter")
        col_filter_btn.clicked.connect(self.pivot_tab.open_column_filter_window)
        control_layout.addWidget(col_filter_btn)

        # Check RM button
        check_rm_btn = QPushButton("Check RM")
        check_rm_btn.clicked.connect(self.pivot_tab.check_rm)
        control_layout.addWidget(check_rm_btn)

        # Clear CRM button
        clear_crm_btn = QPushButton("Clear CRM")
        clear_crm_btn.clicked.connect(self.pivot_tab.clear_inline_crm)
        control_layout.addWidget(clear_crm_btn)

        layout.addWidget(control_frame)

        # Table view
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setStretchLastSection(False)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_view.verticalHeader().setVisible(False)
        self.table_view.setWordWrap(True)
        layout.addWidget(self.table_view)

        # Status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

    def set_model(self, model):
        self.table_view.setModel(model)
        self.pivot_tab.column_widths.clear()
        for col, width in self.pivot_tab.column_widths.items():
            self.table_view.horizontalHeader().resizeSection(col, width)