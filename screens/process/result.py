import sys
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableView,
    QHeaderView, QScrollBar, QComboBox, QLineEdit, QDialog, QFileDialog, QMessageBox, QFrame
)
from PyQt6.QtCore import Qt, QAbstractTableModel, QVariant
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import time
import numpy as np
import os
import platform

# Global stylesheet for consistent UI and white backgrounds
global_style = """
    QWidget {
        background-color: white;
        font: 12px 'Segoe UI';
    }
    QDialog {
        background-color: white;
    }
    QPushButton {
        background-color: #f5f5f5;
        color: black;
        border: 1px solid #ccc;
        padding: 5px;
        font: bold 11px 'Segoe UI';
        min-width: 140px;
        min-height: 30px;
    }
    QPushButton:hover {
        background-color: #ddd;
    }
    QComboBox {
        background-color: white;
        color: black;
        border: 1px solid #ccc;
        padding: 5px;
        font: 12px 'Segoe UI';
    }
    QComboBox QAbstractItemView {
        background-color: white;
        color: black;
        selection-background-color: #ddd;
        selection-color: black;
    }
    QLineEdit {
        background-color: white;
        color: black;
        border: 1px solid #ccc;
        padding: 5px;
        font: 12px 'Segoe UI';
    }
    QTableView {
        background-color: white;
        border: 1px solid #ccc;
        font: 12px 'Segoe UI';
    }
    QHeaderView::section {
        background-color: #d3d3d3;
        font: bold 12px 'Segoe UI';
        border: 1px solid #ccc;
    }
"""

class PandasModel(QAbstractTableModel):
    """Custom model to display pandas DataFrame in QTableView"""
    def __init__(self, data=pd.DataFrame()):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return QVariant()
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        elif role == Qt.ItemDataRole.BackgroundRole:
            return Qt.GlobalColor.lightGray if index.row() % 2 else Qt.GlobalColor.white
        return QVariant()

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])
            return str(self._data.index[section])
        return QVariant()

class FilterDialog(QDialog):
    """Dialog for selecting filter values"""
    def __init__(self, parent, filter_values, filter_field, solution_label_order, element_order, update_callback):
        super().__init__(parent)
        self.setWindowTitle("Filter Pivot Table")
        self.setFixedSize(300, 450)
        self.filter_values = filter_values
        self.filter_field = filter_field
        self.solution_label_order = solution_label_order
        self.element_order = element_order
        self.update_callback = update_callback
        self.setStyleSheet(global_style)  # Apply global stylesheet
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Filter field selection
        layout.addWidget(QLabel("Select Column:", font=QFont("Segoe UI", 12)))
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["Solution Label", "Element"])
        self.filter_combo.setFont(QFont("Segoe UI", 12))
        self.filter_combo.setCurrentText(self.filter_field)  # Fixed: Removed .get()
        self.filter_combo.currentTextChanged.connect(self.update_checkboxes)
        layout.addWidget(self.filter_combo)

        # Table for filter values
        self.filter_table = QTableView()
        self.filter_table.setSelectionMode(QTableView.SelectionMode.NoSelection)
        self.filter_table.setStyleSheet(global_style)
        self.filter_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.filter_table)

        # Scrollbar
        self.filter_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        # Buttons
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(lambda: self.set_all_checkboxes(True))
        layout.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(lambda: self.set_all_checkboxes(False))
        layout.addWidget(deselect_all_btn)

        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.accept)
        layout.addWidget(apply_btn)

        self.update_checkboxes()

    def update_checkboxes(self):
        field = self.filter_combo.currentText()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Value", "Select"])

        unique_values = (
            self.solution_label_order if field == "Solution Label"
            else self.element_order if self.element_order else []
        )
        if field not in self.filter_values:
            self.filter_values[field] = {val: True for val in unique_values}

        for value in unique_values:
            value_item = QStandardItem(value)
            value_item.setEditable(False)
            check_item = QStandardItem()
            check_item.setCheckable(True)
            check_item.setCheckState(
                Qt.CheckState.Checked if self.filter_values[field].get(value, True)
                else Qt.CheckState.Unchecked
            )
            model.appendRow([value_item, check_item])

        self.filter_table.setModel(model)
        self.filter_table.setColumnWidth(0, 150)
        self.filter_table.setColumnWidth(1, 80)
        model.itemChanged.connect(lambda item: self.toggle_filter(item, field))

    def toggle_filter(self, item, field):
        if item.column() == 1:  # Only handle changes in the "Select" column
            value = self.filter_table.model().item(item.row(), 0).text()
            self.filter_values[field][value] = (item.checkState() == Qt.CheckState.Checked)
            self.update_callback()

    def set_all_checkboxes(self, value):
        field = self.filter_combo.currentText()
        if field in self.filter_values:
            for val in self.filter_values[field]:
                self.filter_values[field][val] = value
            self.update_checkboxes()
            self.update_callback()

class ResultsFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.setStyleSheet(global_style)  # Apply global stylesheet
        self.search_var = ""
        self.filter_field = "Solution Label"
        self.filter_values = {}  # Store selected filter values
        self.search_window = None
        self.column_widths = {}  # Cache column widths
        self.last_filtered_data = None  # Cache filtered data
        self._last_cache_key = None  # Cache key for filtering
        self.solution_label_order = None
        self.element_order = None
        self.decimal_places = "1"  # Default to 1 decimal place
        self.setup_ui()

    def setup_ui(self):
        """Setup Results UI with pivot table, scrollbars, and decimal places selection"""
        start_time = time.time()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Button frame
        button_frame = QWidget()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(5)

        # Search button
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.open_search_window)
        button_layout.addWidget(search_button)

        # Filter button
        filter_button = QPushButton("Filter")
        filter_button.clicked.connect(self.open_filter_window)
        button_layout.addWidget(filter_button)

        # Save button
        self.save_button = QPushButton("Save Processed Excel")
        self.save_button.clicked.connect(self.save_processed_excel)
        button_layout.addWidget(self.save_button)

        # Decimal places selection
        decimal_label = QLabel("Decimal Places:")
        decimal_label.setFont(QFont("Segoe UI", 12))
        button_layout.addWidget(decimal_label)
        self.decimal_combo = QComboBox()
        self.decimal_combo.addItems(["0", "1", "2", "3"])
        self.decimal_combo.setCurrentText(self.decimal_places)
        self.decimal_combo.setFont(QFont("Segoe UI", 12))
        self.decimal_combo.setFixedWidth(60)
        self.decimal_combo.currentTextChanged.connect(self.show_processed_data)
        button_layout.addWidget(self.decimal_combo)

        layout.addWidget(button_frame)

        # Table frame
        table_frame = QFrame()
        table_frame.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(0, 0, 0, 0)

        # Table view
        self.processed_table = QTableView()
        self.processed_table.setStyleSheet(global_style)
        self.processed_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.processed_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.processed_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.processed_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        table_layout.addWidget(self.processed_table)

        layout.addWidget(table_frame, stretch=1)
        self.setLayout(layout)

        print(f"ResultsFrame UI setup took {time.time() - start_time:.3f} seconds")

    def format_value(self, x):
        """Format a value to the specified number of decimal places"""
        try:
            decimal_places = int(self.decimal_combo.currentText())
            return f"{float(x):.{decimal_places}f}"
        except (ValueError, TypeError):
            return str(x)

    def get_filtered_data(self):
        """Get filtered pivot table data with caching"""
        start_time = time.time()
        df = self.app.get_data()
        if df is None or df.empty:
            print("No data available in get_filtered_data")
            return None

        required_columns = ['Solution Label', 'Element', 'Corr Con', 'Type']
        if not all(col in df.columns for col in required_columns):
            QMessageBox.warning(self, "Error", f"DataFrame missing required columns: {required_columns}")
            return None

        df_filtered = df[df['Type'].isin(['Samp', 'Sample'])].copy()
        df_filtered = df_filtered[
            (~df_filtered['Solution Label'].isin(self.app.get_excluded_samples())) &
            (~df_filtered['Solution Label'].isin(self.app.get_excluded_volumes())) &
            (~df_filtered['Solution Label'].isin(self.app.get_excluded_dfs()))
        ]

        if df_filtered.empty:
            print("No data after filtering in get_filtered_data")
            return None

        df_filtered['original_index'] = df_filtered.index
        df_filtered['Element'] = df_filtered['Element'].str.split('_').str[0]

        df_filtered['unique_id'] = df_filtered.groupby(['Solution Label', 'Element']).cumcount()

        if self.solution_label_order is None:
            self.solution_label_order = df_filtered['Solution Label'].drop_duplicates().tolist()

        if self.element_order is None:
            self.element_order = df_filtered['Element'].drop_duplicates().tolist()

        value_column = 'Corr Con'
        if value_column not in df_filtered.columns:
            QMessageBox.warning(self, "Error", f"Column '{value_column}' not found in data!")
            return None

        pivot_data = df_filtered.pivot_table(
            index=['Solution Label', 'unique_id'],
            columns='Element',
            values=value_column,
            aggfunc='first'
        ).reset_index()

        pivot_data = pivot_data.merge(
            df_filtered[['Solution Label', 'unique_id', 'original_index']].drop_duplicates(),
            on=['Solution Label', 'unique_id'],
            how='left'
        ).sort_values('original_index').drop(columns=['unique_id'])

        columns_to_keep = ['Solution Label'] + [col for col in self.element_order if col in pivot_data.columns]
        pivot_data = pivot_data[columns_to_keep]

        search_text = self.search_var.lower()
        filter_field = self.filter_field
        selected_values = [k for k, v in self.filter_values.get(filter_field, {}).items() if v] if filter_field else []

        cache_key = (search_text, filter_field, tuple(sorted(selected_values)))
        if cache_key == self._last_cache_key and self.last_filtered_data is not None:
            return self.last_filtered_data

        if search_text:
            mask = pivot_data.astype(str).apply(lambda x: x.str.lower().str.contains(search_text, na=False)).any(axis=1)
            pivot_data = pivot_data[mask]

        if filter_field and selected_values:
            if filter_field == 'Solution Label':
                selected_order = [x for x in self.solution_label_order if x in selected_values and x in pivot_data['Solution Label'].values]
                pivot_data = pivot_data[pivot_data['Solution Label'].isin(selected_values)]
                pivot_data['Solution Label'] = pd.Categorical(
                    pivot_data['Solution Label'],
                    categories=selected_order,
                    ordered=True
                )
                pivot_data = pivot_data.sort_values('Solution Label').reset_index(drop=True)
            elif filter_field == 'Element':
                columns_to_keep = ['Solution Label'] + [col for col in self.element_order if col in selected_values and col in pivot_data.columns]
                pivot_data = pivot_data[columns_to_keep]

        pivot_data = pivot_data.drop_duplicates().reset_index(drop=True)
        self.last_filtered_data = pivot_data
        self._last_cache_key = cache_key
        return pivot_data

    def show_processed_data(self):
        """Show the final processed pivot table"""
        start_time = time.time()
        df = self.get_filtered_data()
        model = QStandardItemModel()

        if df is None:
            model.setHorizontalHeaderLabels(["Status"])
            model.appendRow([QStandardItem("No data loaded")])
            self.processed_table.setModel(model)
            self.processed_table.setColumnWidth(0, 150)
            self.column_widths = {"Status": 150}
            print(f"Show processed data (no data) took {time.time() - start_time:.3f} seconds")
            return

        # Set headers
        columns = list(df.columns)
        model.setHorizontalHeaderLabels(columns)
        self.column_widths = {}

        # Calculate column widths
        for col_idx, col in enumerate(columns):
            max_width = max(
                [len(str(col))] + [len(self.format_value(x)) for x in df[col].dropna()],
                default=10
            )
            pixel_width = min(max_width * 10, 150)
            self.column_widths[col] = pixel_width
            self.processed_table.setColumnWidth(col_idx, pixel_width)

        # Populate table
        for row_idx, row in df.iterrows():
            items = [QStandardItem(self.format_value(x)) for x in row]
            for item in items:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            model.appendRow(items)

        self.processed_table.setModel(model)
        print(f"Show processed pivot data took {time.time() - start_time:.3f} seconds")

    def open_search_window(self):
        """Open a dialog for search input"""
        start_time = time.time()
        if self.app.get_data() is None:
            QMessageBox.warning(self, "Warning", "No data to search!")
            return

        if self.search_window and self.search_window.isVisible():
            self.search_window.close()

        self.search_window = QDialog(self)
        self.search_window.setWindowTitle("Search Pivot Table")
        self.search_window.setFixedSize(300, 150)
        self.search_window.setStyleSheet(global_style)  # Apply global stylesheet
        layout = QVBoxLayout(self.search_window)

        layout.addWidget(QLabel("Search:", font=QFont("Segoe UI", 12)))
        search_entry = QLineEdit()
        search_entry.setFont(QFont("Segoe UI", 12))
        search_entry.textChanged.connect(lambda text: setattr(self, 'search_var', text))
        layout.addWidget(search_entry)

        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self.show_processed_data)
        layout.addWidget(search_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.search_window.close)
        layout.addWidget(close_btn)

        self.search_window.show()
        search_entry.setFocus()
        print(f"Open search window took {time.time() - start_time:.3f} seconds")

    def open_filter_window(self):
        """Open a dialog to select filter values"""
        start_time = time.time()
        df = self.app.get_data()
        if df is None:
            QMessageBox.warning(self, "Warning", "No data to filter!")
            return

        dialog = FilterDialog(
            self, self.filter_values, self.filter_field,
            self.solution_label_order, self.element_order,
            self.show_processed_data
        )
        dialog.accepted.connect(self.show_processed_data)
        dialog.exec()
        self.filter_field = dialog.filter_combo.currentText()
        print(f"Open filter window took {time.time() - start_time:.3f} seconds")

    def save_processed_excel(self):
        """Save processed pivot table to Excel with styles"""
        start_time = time.time()
        df = self.get_filtered_data()
        if df is None:
            QMessageBox.warning(self, "Warning", "No data to save!")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed Pivot Table"

                header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                first_column_fill = PatternFill(start_color="FFF5E4", end_color="FFF5E4", fill_type="solid")
                odd_row_fill = PatternFill(start_color="f5f5f5", end_color="f5f5f5", fill_type="solid")
                even_row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                header_font = Font(name="Segoe UI", size=12, bold=True)
                cell_font = Font(name="Segoe UI", size=12)
                cell_alignment = Alignment(horizontal="center", vertical="center")
                thin_border = Border(
                    left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin")
                )

                headers = list(df.columns)
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.value = header
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = cell_alignment
                    cell.border = thin_border
                    ws.column_dimensions[get_column_letter(col_idx)].width = 15

                decimal_places = int(self.decimal_combo.currentText())
                for row_idx, (_, row) in enumerate(df.iterrows(), 2):
                    fill = even_row_fill if (row_idx - 1) % 2 == 0 else odd_row_fill
                    for col_idx, value in enumerate(row, 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        try:
                            cell.value = round(float(value), decimal_places)
                            cell.number_format = f"0.{'0' * decimal_places}"
                        except (ValueError, TypeError):
                            cell.value = value
                        cell.font = cell_font
                        cell.alignment = cell_alignment
                        cell.border = thin_border
                        if col_idx == 1:
                            cell.fill = first_column_fill
                        else:
                            cell.fill = fill

                wb.save(file_path)
                QMessageBox.information(self, "Success", "Processed pivot table saved successfully!")
                print(f"Save processed excel took {time.time() - start_time:.3f} seconds")

                if QMessageBox.question(self, "Open File", "Would you like to open the saved Excel file?") == QMessageBox.StandardButton.Yes:
                    try:
                        system = platform.system()
                        if system == "Windows":
                            os.startfile(file_path)
                        elif system == "Darwin":
                            os.system(f"open {file_path}")
                        else:
                            os.system(f"xdg-open {file_path}")
                        print(f"Opened file: {file_path}")
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")
                        print(f"Failed to open file: {str(e)}")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save: {str(e)}")

    def reset_cache(self):
        """Reset cached data, orders, filters, and search"""
        self.last_filtered_data = None
        self._last_cache_key = None
        self.solution_label_order = None
        self.element_order = None
        self.column_widths = {}
        self.filter_values = {}
        self.search_var = ""
        print("ResultsFrame cache, orders, filters, and search reset")