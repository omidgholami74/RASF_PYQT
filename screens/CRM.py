import sys
import os
import platform
import pandas as pd
import sqlite3
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QTableView, 
                             QFrame, QScrollArea, QGridLayout, QDialog, QMessageBox, QHeaderView,
                             QLineEdit, QCheckBox, QScrollBar)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel
from PyQt6.QtGui import QFont, QStandardItemModel, QStandardItem, QColor, QPalette
from openpyxl import load_workbook
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Global stylesheet (defined above)
global_style = """
    QWidget {
        background-color: white;
        font: 12px 'Segoe UI';
    }
    QDialog {
        background-color: white;
        border: 1px solid #cccccc;
        border-radius: 4px;
    }
    QPushButton {
        background-color: #f5f5f5;
        color: black;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 8px 12px;
        font: bold 11px 'Segoe UI';
        min-width: 110px;
        min-height: 32px;
    }
    QPushButton:hover {
        background-color: #e0e0e0;
        border: 1px solid #aaa;
    }
    QPushButton:pressed {
        background-color: #d0d0d0;
    }
    QComboBox {
        background-color: white;
        color: black;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 8px;
        font: 11px 'Segoe UI';
    }
    QComboBox:hover {
        border: 1px solid #aaa;
    }
    QComboBox QAbstractItemView {
        background-color: white;
        color: black;
        selection-background-color: #e0e0e0;
        selection-color: black;
        border: 1px solid #ccc;
        border-radius: 4px;
    }
    QLineEdit {
        background-color: white;
        color: black;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 8px;
        font: 11px 'Segoe UI';
    }
    QLineEdit:focus {
        border: 1px solid #0078d4;
    }
    QTableView {
        background-color: white;
        alternate-background-color: #f9f9f9;
        border: 1px solid #ccc;
        font: 11px 'Segoe UI';
        selection-background-color: #ADD8E6;
        selection-color: black;
        gridline-color: #e0e0e0;
    }
    QHeaderView::section {
        background-color: #e0e0e0;
        color: black;
        border: 1px solid #ccc;
        padding: 6px;
        font: bold 11px 'Segoe UI';
    }
    QHeaderView::section:hover {
        background-color: #d0d0d0;
    }
    QScrollArea {
        background-color: white;
        border: none;
    }
    QScrollBar:vertical, QScrollBar:horizontal {
        background: #f0f0f0;
        border: 1px solid #ccc;
        border-radius: 3px;
    }
    QScrollBar::handle {
        background: #aaaaaa;
        border-radius: 3px;
    }
    QScrollBar::handle:hover {
        background: #888888;
    }
    QLabel {
        color: black;
        font: 12px 'Segoe UI';
    }
"""

class CRMTableModel(QAbstractTableModel):
    """Custom model for QTableView to display CRM data efficiently."""
    def __init__(self, crm_tab, df=None, decimal_places=2):
        super().__init__()
        self.crm_tab = crm_tab
        self._df = df if df is not None else pd.DataFrame()
        self.decimal_places = decimal_places
        self.column_widths = {}

    def set_data(self, df):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return self._df.shape[0]

    def columnCount(self, parent=QModelIndex()):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < self._df.shape[0] and 0 <= index.column() < self._df.shape[1]):
            return None

        value = self._df.iloc[index.row(), index.column()]
        col_name = self._df.columns[index.column()]

        if role == Qt.ItemDataRole.DisplayRole:
            dec = self.decimal_places
            if col_name != 'CRM ID' and pd.notna(value):
                try:
                    return f"{float(value):.{dec}f}"
                except (ValueError, TypeError):
                    return "" if pd.isna(value) else str(value)
            return str(value) if pd.notna(value) else ""
        
        elif role == Qt.ItemDataRole.BackgroundRole:
            return QColor("#f9f9f9") if index.row() % 2 == 0 else QColor("white")
        
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignLeft if col_name == 'CRM ID' else Qt.AlignmentFlag.AlignCenter

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._df.columns[section])
            return str(section + 1)
        return None

class CRMTab(QWidget):
    """CRMTab for managing CRM data with SQLite and pivot table display."""
    def __init__(self, app, parent_frame):
        super().__init__(parent_frame)
        self.app = app
        self.parent_frame = parent_frame
        self.conn = None
        self.crm_data = None
        self.pivot_data = None
        self.table_view = None
        self.search_var = QLineEdit()
        self.filter_var = QComboBox()
        self.decimal_places = QComboBox()
        self.column_widths = {}
        self.sort_column = None
        self.sort_reverse = False
        self.ui_initialized = False
        self.setStyleSheet(global_style)
        self.setup_ui()
        self.init_db()

    def setup_ui(self):
        """Setup UI with controls and placeholder."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(1, 1, 1,1)  # Increased margins for breathing room
        main_layout.setSpacing(1)  # Slightly larger spacing for clarity

        # Control frame
        control_frame = QFrame()
        control_layout = QHBoxLayout(control_frame)
        control_layout.setContentsMargins(6, 6, 6, 6)  # Optimized margins
        control_layout.setSpacing(12)  # Consistent spacing

        # Analysis Method Filter
        control_layout.addWidget(QLabel("Analysis Method:", font=QFont("Segoe UI", 11)))
        self.filter_var.setFixedWidth(180)
        self.filter_var.setToolTip("Filter by analysis method")
        self.filter_var.currentTextChanged.connect(self.update_display)
        control_layout.addWidget(self.filter_var)

        # Decimal Places
        control_layout.addWidget(QLabel("Decimals:", font=QFont("Segoe UI", 11)))
        self.decimal_places.addItems(["0", "1", "2", "3"])
        self.decimal_places.setCurrentText("2")
        self.decimal_places.setFixedWidth(80)
        self.decimal_places.setToolTip("Set decimal places for numeric values")
        self.decimal_places.currentTextChanged.connect(self.update_display)
        control_layout.addWidget(self.decimal_places)

        # Search Field
        self.search_var.setPlaceholderText("Search by any field...")
        self.search_var.setFixedWidth(180)
        self.search_var.setToolTip("Search across all columns")
        self.search_var.textChanged.connect(self.update_display)
        control_layout.addWidget(self.search_var)

        # CRUD Buttons
        add_btn = QPushButton("Add Record")
        add_btn.setToolTip("Add a new CRM record")
        add_btn.clicked.connect(self.open_add_window)
        control_layout.addWidget(add_btn)

        edit_btn = QPushButton("Edit Record")
        edit_btn.setToolTip("Edit the selected record")
        edit_btn.clicked.connect(self.open_edit_window)
        control_layout.addWidget(edit_btn)

        delete_btn = QPushButton("Delete Selected")
        delete_btn.setToolTip("Delete the selected record")
        delete_btn.clicked.connect(self.delete_selected)
        control_layout.addWidget(delete_btn)

        control_layout.addStretch()
        main_layout.addWidget(control_frame)

        # Placeholder label
        self.placeholder_label = QLabel("Click 'Display' to load data.", font=QFont("Segoe UI", 12))
        self.placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder_label.setStyleSheet("color: #666666;")
        main_layout.addWidget(self.placeholder_label)

        # Apply styles
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor("white"))
        self.setPalette(palette)

    def setup_full_ui(self):
        """Setup full UI with QTableView and scrollbars."""
        if self.ui_initialized:
            return

        # Remove placeholder
        if self.placeholder_label:
            self.placeholder_label.deleteLater()

        # Table container
        table_container = QFrame()
        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)  # No margins for table
        table_layout.setSpacing(0)

        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table_view.setSortingEnabled(True)
        self.table_view.setFont(QFont("Segoe UI", 11))
        self.table_view.horizontalHeader().setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_view.horizontalHeader().setMinimumSectionSize(100)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.verticalHeader().setVisible(False)
        self.table_view.setEnabled(True)
        self.table_view.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.table_view.setStyleSheet(global_style)
        self.table_view.setShowGrid(True)  # Ensure grid lines are visible
        table_layout.addWidget(self.table_view)

        # Horizontal scrollbar
        hsb = QScrollBar(Qt.Orientation.Horizontal)
        self.table_view.setHorizontalScrollBar(hsb)
        table_layout.addWidget(hsb)

        main_layout = self.layout()
        main_layout.addWidget(table_container)

        self.ui_initialized = True

    def init_db(self):
        """Initialize SQLite database."""
        try:
            self.conn = sqlite3.connect("crm_data.db")
            logger.info("Connected to SQLite database")
        except Exception as e:
            logger.error(f"Failed to connect to SQLite database: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to connect to database:\n{str(e)}")

    def load_and_display(self):
        """Load data from SQLite or Excel and display pivot table."""
        try:
            if self.conn is None:
                self.init_db()
            cursor = self.conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='crm'")
            if not cursor.fetchone():
                excel_data = pd.read_excel("data crm.xlsx")
                excel_data.to_sql("crm", self.conn, if_exists="replace", index=False)
                logger.info("Excel data loaded into SQLite database")
                self.conn.commit()

            self.setup_full_ui()
            self.update_filter_options()
            self.update_display()
        except Exception as e:
            logger.error(f"Failed to load CRM data: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to load CRM data:\n{str(e)}")

    def update_filter_options(self):
        """Update the Analysis Method filter dropdown with unique values."""
        if self.conn is None:
            self.filter_var.addItem("All")
            self.filter_var.setCurrentText("All")
            return

        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT DISTINCT [Analysis Method] FROM crm WHERE [Analysis Method] IS NOT NULL")
            unique_methods = ["All"] + sorted([row[0] for row in cursor.fetchall()])
            self.filter_var.clear()
            self.filter_var.addItems(unique_methods)
            if self.filter_var.currentText() not in unique_methods:
                self.filter_var.setCurrentText("All")
        except Exception as e:
            logger.error(f"Failed to update filter options: {str(e)}")
            self.filter_var.clear()
            self.filter_var.addItem("All")
            self.filter_var.setCurrentText("All")

    def update_display(self, event=None):
        """Update QTableView display with pivot table data."""
        if self.conn is None:
            self._set_status_table("No data loaded")
            return

        try:
            cursor = self.conn.cursor()
            query = "SELECT * FROM crm"
            conditions = []
            params = []
            search_text = self.search_var.text().lower()
            filter_method = self.filter_var.currentText()

            cursor.execute("PRAGMA table_info(crm)")
            columns = [info[1] for info in cursor.fetchall()]

            if search_text:
                search_conditions = [f"LOWER([{col}]) LIKE ?" for col in columns]
                conditions.append("(" + " OR ".join(search_conditions) + ")")
                params.extend([f"%{search_text}%" for _ in columns])

            if filter_method != "All" and "Analysis Method" in columns:
                conditions.append("[Analysis Method] = ?")
                params.append(filter_method)

            if conditions:
                query += " WHERE " + " AND ".join(conditions)

            df = pd.read_sql_query(query, self.conn, params=params)
            if df.empty:
                self._set_status_table("No data after filtering")
                return

            # Process Element column to keep only the part after the comma
            if 'Element' in df.columns:
                df['Element'] = df['Element'].apply(lambda x: x.split(', ')[-1].strip() if isinstance(x, str) and ', ' in x else x)

            # Create pivot table
            required_columns = {'CRM ID', 'Element', 'Sort Grade'}
            if not required_columns.issubset(df.columns):
                self._set_status_table("Missing required columns")
                return

            # Aggregate to avoid duplicates
            aggregated_df = df.groupby(['CRM ID', 'Element'])['Sort Grade'].first().reset_index()

            # Pivot table
            pivot_df = pd.pivot_table(
                aggregated_df,
                index='CRM ID',
                columns='Element',
                values='Sort Grade',
                aggfunc='first'
            ).reset_index()

            self.pivot_data = pivot_df
            pivot_df.to_csv('pivot_crm.csv')
            model = CRMTableModel(self, pivot_df, decimal_places=int(self.decimal_places.currentText()))
            self.table_view.setModel(model)
            
            # Set column widths
            for col_idx, col in enumerate(pivot_df.columns):
                max_width = max([len(str(x)) for x in pivot_df[col].dropna()] + [len(str(col))], default=10)
                pixel_width = min(max_width * 8, 160)
                self.column_widths[col] = pixel_width
                self.table_view.setColumnWidth(col_idx, pixel_width)

        except Exception as e:
            logger.error(f"Failed to update display: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to update display:\n{str(e)}")

    def _set_status_table(self, text):
        """Set status message in the table."""
        self.table_view.setModel(None)
        self.status_label = QLabel(text, font=QFont("Segoe UI", 12))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout = self.layout()
        layout.addWidget(self.status_label)

    def open_search_window(self):
        """Open a window for search input."""
        if self.conn is None:
            QMessageBox.warning(self, "Warning", "No data to search!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Search CRM Table")
        dialog.setGeometry(200, 200, 300, 150)
        dialog.setModal(True)
        dialog.setStyleSheet(global_style)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        layout.addWidget(QLabel("Search:", font=QFont("Segoe UI", 11)))
        search_entry = QLineEdit()
        search_entry.setText(self.search_var.text())
        search_entry.setFont(QFont("Segoe UI", 11))
        search_entry.setToolTip("Enter search term")
        search_entry.textChanged.connect(lambda text: self.search_var.setText(text))
        layout.addWidget(search_entry)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        search_btn = QPushButton("Search")
        search_btn.setFixedSize(120, 34)
        search_btn.setToolTip("Apply search filter")
        search_btn.clicked.connect(self.update_display)
        button_layout.addWidget(search_btn)

        close_btn = QPushButton("Close")
        close_btn.setFixedSize(120, 34)
        close_btn.setToolTip("Close this dialog")
        close_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)
        dialog.exec()

    def open_filter_window(self):
        """Open a window for Analysis Method filter."""
        if self.conn is None:
            QMessageBox.warning(self, "Warning", "No data to filter!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Filter Analysis Method")
        dialog.setGeometry(200, 200, 300, 150)
        dialog.setModal(True)
        dialog.setStyleSheet(global_style)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        layout.addWidget(QLabel("Select Analysis Method:", font=QFont("Segoe UI", 11)))
        filter_combo = QComboBox()
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT [Analysis Method] FROM crm WHERE [Analysis Method] IS NOT NULL")
        unique_methods = ["All"] + sorted([row[0] for row in cursor.fetchall()])
        filter_combo.addItems(unique_methods)
        filter_combo.setFont(QFont("Segoe UI", 11))
        filter_combo.setCurrentText(self.filter_var.currentText())
        filter_combo.setToolTip("Select an analysis method to filter")
        filter_combo.currentTextChanged.connect(lambda text: self.filter_var.setCurrentText(text))
        layout.addWidget(filter_combo)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        apply_btn = QPushButton("Apply")
        apply_btn.setFixedSize(120, 34)
        apply_btn.setToolTip("Apply the selected filter")
        apply_btn.clicked.connect(self.update_display)
        button_layout.addWidget(apply_btn)

        close_btn = QPushButton("Close")
        close_btn.setFixedSize(120, 34)
        close_btn.setToolTip("Close this dialog")
        close_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)
        dialog.exec()

    def open_add_window(self):
        """Open a window to add a new record."""
        if self.conn is None:
            QMessageBox.warning(self, "Warning", "No database connection!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Add New CRM Record")
        dialog.setGeometry(200, 200, 380, 540)
        dialog.setModal(True)
        dialog.setStyleSheet(global_style)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(14)

        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(crm)")
        columns = [info[1] for info in cursor.fetchall()]
        entry_vars = {col: QLineEdit() for col in columns}

        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        grid_layout = QGridLayout(scroll_widget)
        grid_layout.setContentsMargins(6, 6, 6, 6)
        grid_layout.setSpacing(10)
        for i, col in enumerate(columns):
            label = QLabel(f"{col}:", font=QFont("Segoe UI", 11))
            label.setFixedWidth(160)
            grid_layout.addWidget(label, i, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            entry = entry_vars[col]
            entry.setFixedWidth(180)
            entry.setToolTip(f"Enter value for {col}")
            grid_layout.addWidget(entry, i, 1)

        scroll_widget.setLayout(grid_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        save_btn = QPushButton("Save")
        save_btn.setFixedSize(120, 34)
        save_btn.setToolTip("Save the new record")
        save_btn.clicked.connect(lambda: self.save_record(entry_vars, dialog))
        button_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedSize(120, 34)
        cancel_btn.setToolTip("Cancel and close")
        cancel_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)
        dialog.exec()

    def save_record(self, entry_vars, dialog):
        try:
            cursor = self.conn.cursor()
            cursor.execute("PRAGMA table_info(crm)")
            columns = [info[1] for info in cursor.fetchall()]
            values = [entry_vars[col].text() or None for col in columns]
            query = f"INSERT INTO crm ({', '.join(f'[{col}]' for col in columns)}) VALUES ({', '.join(['?'] * len(columns))})"
            cursor.execute(query, values)
            self.conn.commit()
            logger.info("Added new record to crm table")
            self.update_display()
            QMessageBox.information(dialog, "Success", "Record added successfully!")
            dialog.accept()
        except Exception as e:
            logger.error(f"Failed to add record: {str(e)}")
            QMessageBox.warning(dialog, "Error", f"Failed to add record:\n{str(e)}")

    def open_edit_window(self):
        """Open a window to edit the selected record."""
        if self.conn is None:
            QMessageBox.warning(self, "Warning", "No database connection!")
            return

        selected = self.table_view.selectionModel().selectedRows()
        logger.debug(f"Selected rows: {len(selected)}")
        if not selected:
            QMessageBox.warning(self, "Warning", "Please select a record to edit!")
            return
        if len(selected) > 1:
            QMessageBox.warning(self, "Warning", "Please select only one record to edit!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit CRM Record")
        dialog.setGeometry(200, 200, 380, 540)
        dialog.setModal(True)
        dialog.setStyleSheet(global_style)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(14)

        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(crm)")
        columns = [info[1] for info in cursor.fetchall()]
        row = selected[0].row()
        values = [self.table_view.model().data(self.table_view.model().index(row, i)) for i in range(self.table_view.model().columnCount())]
        entry_vars = {col: QLineEdit(str(values[i]) if values[i] is not None else "") for i, col in enumerate(columns)}

        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        grid_layout = QGridLayout(scroll_widget)
        grid_layout.setContentsMargins(6, 6, 6, 6)
        grid_layout.setSpacing(10)
        for i, col in enumerate(columns):
            label = QLabel(f"{col}:", font=QFont("Segoe UI", 11))
            label.setFixedWidth(160)
            grid_layout.addWidget(label, i, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            entry = entry_vars[col]
            entry.setFixedWidth(180)
            entry.setToolTip(f"Edit value for {col}")
            grid_layout.addWidget(entry, i, 1)

        scroll_widget.setLayout(grid_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        save_btn = QPushButton("Save")
        save_btn.setFixedSize(120, 34)
        save_btn.setToolTip("Save changes")
        save_btn.clicked.connect(lambda: self.save_edit(entry_vars, values[0], dialog))
        button_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedSize(120, 34)
        cancel_btn.setToolTip("Cancel and close")
        cancel_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)
        dialog.exec()

    def save_edit(self, entry_vars, id_value, dialog):
        try:
            cursor = self.conn.cursor()
            cursor.execute("PRAGMA table_info(crm)")
            columns = [info[1] for info in cursor.fetchall()]
            set_clause = ", ".join(f"[{col}] = ?" for col in columns[1:])
            query = f"UPDATE crm SET {set_clause} WHERE [{columns[0]}] = ?"
            update_values = [entry_vars[col].text() or None for col in columns[1:]] + [id_value]
            cursor.execute(query, update_values)
            self.conn.commit()
            logger.info(f"Updated record with {columns[0]} = {id_value}")
            self.update_display()
            QMessageBox.information(dialog, "Success", "Record updated successfully!")
            dialog.accept()
        except Exception as e:
            logger.error(f"Failed to update record: {str(e)}")
            QMessageBox.warning(dialog, "Error", f"Failed to update record:\n{str(e)}")

    def delete_selected(self):
        """Delete selected records from the database."""
        if self.conn is None:
            QMessageBox.warning(self, "Warning", "No database connection!")
            return

        selected = self.table_view.selectionModel().selectedRows()
        if not selected:
            QMessageBox.warning(self, "Warning", "Please select at least one record to delete!")
            return

        if not QMessageBox.question(self, "Confirm Delete", "Are you sure you want to delete the selected records?") == QMessageBox.StandardButton.Yes:
            return

        try:
            cursor = self.conn.cursor()
            cursor.execute("PRAGMA table_info(crm)")
            columns = [info[1] for info in cursor.fetchall()]
            id_col = columns[0]
            for row in selected:
                id_value = self.table_view.model().data(self.table_view.model().index(row.row(), 0))
                cursor.execute(f"DELETE FROM crm WHERE [{id_col}] = ?", (id_value,))
                logger.info(f"Deleted record with {id_col} = {id_value}")
            self.conn.commit()
            self.update_display()
            QMessageBox.information(self, "Success", "Selected records deleted successfully!")
        except Exception as e:
            logger.error(f"Failed to delete records: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to delete records:\n{str(e)}")

    def reset_cache(self):
        """Reset cache and close database connection."""
        if self.conn:
            self.conn.close()
            logger.info("SQLite database connection closed")
        self.conn = None
        self.crm_data = None
        self.pivot_data = None
        self.column_widths = {}
        self.search_var.clear()
        self.filter_var.setCurrentText("All")
        self.ui_initialized = False

    def __del__(self):
        """Ensure database connection is closed when object is destroyed."""
        if self.conn:
            self.conn.close()
            logger.info("SQLite database connection closed in destructor")