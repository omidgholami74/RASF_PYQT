from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QComboBox, QLabel, QFrame, QLineEdit, QCheckBox, QDialog, QHeaderView, QTableView, QScrollArea)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt
from .freeze_table_widget import FreezeTableWidget
from .pivot_table_model import PivotTableModel
from .pivot_plot_dialog import PivotPlotDialog
from .crm_manager import CRMManager
from .pivot_creator import PivotCreator
from .pivot_exporter import PivotExporter
from .oxide_factors import oxide_factors
import pandas as pd
import logging
import numpy as np

class FilterDialog(QDialog):
    """Dialog for filtering rows or columns with checkboxes."""
    def __init__(self, parent, title, is_row_filter=True):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.parent = parent
        self.is_row_filter = is_row_filter
        self.checkboxes = {}
        self.layout = QVBoxLayout(self)

        # Scroll area for checkboxes
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        self.layout.addWidget(scroll)

        self.populate_checkboxes(scroll_layout)

        # Buttons
        buttons = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all)
        buttons.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(self.deselect_all)
        buttons.addWidget(deselect_all_btn)

        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.apply_and_close)
        buttons.addWidget(ok_btn)

        self.layout.addLayout(buttons)

    def populate_checkboxes(self, layout):
        if self.parent.pivot_data is None:
            return

        if self.is_row_filter:
            field = 'Solution Label'  # Assuming row filter is on Solution Label
            unique_values = sorted(self.parent.pivot_data[field].unique())
            filter_values = self.parent.row_filter_values.setdefault(field, {})
        else:
            field = 'Element'
            if self.parent.use_oxide_var.isChecked():
                unique_values = sorted([oxide_factors[el][0] for el in oxide_factors if oxide_factors[el][0] in self.parent.pivot_data.columns])
            else:
                unique_values = sorted([col for col in self.parent.pivot_data.columns if col != 'Solution Label'])
            filter_values = self.parent.column_filter_values.setdefault(field, {})

        for val in unique_values:
            if val not in filter_values:
                filter_values[val] = True
            cb = QCheckBox(str(val))
            cb.setChecked(filter_values[val])
            cb.stateChanged.connect(lambda state, v=val: self.update_filter(v, state))
            self.checkboxes[val] = cb
            layout.addWidget(cb)

    def update_filter(self, value, state):
        field = 'Solution Label' if self.is_row_filter else 'Element'
        filter_values = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        filter_values[field][value] = bool(state == Qt.CheckState.Checked.value)

    def select_all(self):
        for cb in self.checkboxes.values():
            cb.setChecked(True)

    def deselect_all(self):
        for cb in self.checkboxes.values():
            cb.setChecked(False)

    def apply_and_close(self):
        self.parent.update_pivot_display()
        self.accept()

class PivotTab(QWidget):
    """PivotTab with inline CRM rows, difference coloring, and plot visualization."""
    def __init__(self, app, parent_frame):
        super().__init__(parent_frame)
        self.logger = logging.getLogger(__name__)
        self.app = app
        self.parent_frame = parent_frame
        self.pivot_data = None
        self.solution_label_order = None
        self.element_order = None
        self.row_filter_values = {}
        self.column_filter_values = {}
        self.original_df = None
        self.column_widths = {}
        self.cached_formatted = {}
        self.current_view_df = None
        self._inline_crm_rows = {}
        self._inline_crm_rows_display = {}
        self._crm_inserted_for_index = set()
        self.current_plot_dialog = None
        self.search_var = QLineEdit()
        self.row_filter_field = QComboBox()
        self.column_filter_field = QComboBox()
        self.element_selector = QComboBox()
        self.decimal_places = QComboBox()
        self.use_int_var = QCheckBox("Use Int")
        self.use_oxide_var = QCheckBox("Use Oxide")
        self.diff_min = QLineEdit("-12")
        self.diff_max = QLineEdit("12")
        self.show_check_crm = QCheckBox("Show Check CRM", checked=True)
        self.show_pivot_crm = QCheckBox("Show Pivot CRM", checked=True)
        self.show_middle = QCheckBox("Show Middle", checked=True)
        self.show_diff = QCheckBox("Show Diff", checked=True)
        self.show_range = QCheckBox("Show Range", checked=True)
        self.max_correction_percent = QLineEdit("32")
        self.included_crms = {}
        self.crm_manager = CRMManager(self)
        self.pivot_creator = PivotCreator(self)
        self.pivot_exporter = PivotExporter(self)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        control_frame = QFrame()
        control_layout = QHBoxLayout(control_frame)
        
        control_layout.addWidget(QLabel("Decimal Places:"))
        self.decimal_places.addItems(["0", "1", "2", "3"])
        self.decimal_places.setCurrentText("1")
        self.decimal_places.currentTextChanged.connect(self.update_pivot_display)
        control_layout.addWidget(self.decimal_places)
        
        self.use_int_var.toggled.connect(self.pivot_creator.create_pivot)
        control_layout.addWidget(self.use_int_var)
        
        self.use_oxide_var.toggled.connect(self.pivot_creator.create_pivot)
        control_layout.addWidget(self.use_oxide_var)
        
        control_layout.addWidget(QLabel("Diff Range (%):"))
        self.diff_min.textChanged.connect(self.validate_diff_range)
        control_layout.addWidget(self.diff_min)
        control_layout.addWidget(QLabel("to"))
        self.diff_max.textChanged.connect(self.validate_diff_range)
        control_layout.addWidget(self.diff_max)
        
        control_layout.addWidget(QLabel("Select Element:"))
        self.element_selector.currentTextChanged.connect(self.show_element_plot)
        control_layout.addWidget(self.element_selector)
        
        plot_btn = QPushButton("Show Plot")
        plot_btn.clicked.connect(self.show_element_plot)
        control_layout.addWidget(plot_btn)
        
        self.search_var.setPlaceholderText("Search...")
        self.search_var.textChanged.connect(self.update_pivot_display)
        control_layout.addWidget(self.search_var)
        
        row_filter_btn = QPushButton("Row Filter")
        row_filter_btn.clicked.connect(self.open_row_filter_window)
        control_layout.addWidget(row_filter_btn)
        
        col_filter_btn = QPushButton("Column Filter")
        col_filter_btn.clicked.connect(self.open_column_filter_window)
        control_layout.addWidget(col_filter_btn)
        
        check_rm_btn = QPushButton("Check RM")
        check_rm_btn.clicked.connect(self.crm_manager.check_rm)
        control_layout.addWidget(check_rm_btn)
        
        clear_crm_btn = QPushButton("Clear CRM")
        clear_crm_btn.clicked.connect(self.clear_inline_crm)
        control_layout.addWidget(clear_crm_btn)
        
        export_btn = QPushButton("Export")
        export_btn.clicked.connect(self.pivot_exporter.export_pivot)
        control_layout.addWidget(export_btn)
        
        layout.addWidget(control_frame)
        
        self.table_view = FreezeTableWidget(PivotTableModel(self))
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_view.doubleClicked.connect(self.on_cell_double_click)
        layout.addWidget(self.table_view)
        
        self.status_label = QLabel("Pivot table will be displayed here.")
        self.status_label.setFont(QFont("Segoe UI", 14))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)

    def validate_diff_range(self):
        try:
            min_val = float(self.diff_min.text())
            max_val = float(self.diff_max.text())
            if min_val > max_val:
                self.diff_min.setText(str(max_val - 1))
            self.update_pivot_display()
        except ValueError:
            pass

    def update_pivot_display(self):
        if self.pivot_data is None or self.pivot_data.empty:
            self.status_label.setText("No data loaded")
            self.table_view.setModel(None)
            self.table_view.frozenTableView.setModel(None)
            return

        df = self.pivot_data.copy()
        print("Pivot data in update_pivot_display:", df)  # Debug

        s = self.search_var.text().strip().lower()
        if s:
            mask = df.apply(lambda r: r.astype(str).str.lower().str.contains(s, na=False).any(), axis=1)
            df = df[mask]

        # Apply row filters
        for field, values in self.row_filter_values.items():
            if field in df.columns:
                selected = [k for k, v in values.items() if v]
                if selected:
                    df = df[df[field].isin(selected)]

        # Apply column filters
        selected_cols = ['Solution Label']
        if self.use_oxide_var.isChecked():
            for field, values in self.column_filter_values.items():
                if field == 'Element':
                    selected_cols.extend([
                        oxide_factors[el][0] for el, v in values.items()
                        if v and el in oxide_factors and oxide_factors[el][0] in df.columns
                    ])
        else:
            for field, values in self.column_filter_values.items():
                if field == 'Element':
                    selected_cols.extend([k for k, v in values.items() if v and k in df.columns])

        if len(selected_cols) > 1:
            df = df[selected_cols]

        df = df.reset_index(drop=True)
        self.current_view_df = df
        print("Current view data:", df)  # Debug

        # Rebuild CRM display for current columns without losing original CRM data
        self._inline_crm_rows_display = self.crm_manager._build_crm_row_lists_for_columns(list(df.columns))
        print("CRM rows display:", self._inline_crm_rows_display)  # Debug

        crm_rows = []
        for sol_label in df['Solution Label']:
            if sol_label in self._inline_crm_rows_display:
                crm_rows.append((sol_label, self._inline_crm_rows_display[sol_label]))

        model = PivotTableModel(self, df, crm_rows)
        self.table_view.setModel(model)
        self.table_view.frozenTableView.setModel(model)
        self.table_view.model().layoutChanged.emit()
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for col, width in self.column_widths.items():
            if col < len(df.columns):
                self.table_view.horizontalHeader().resizeSection(col, width)
        self.status_label.setText("Data loaded successfully")
        self.table_view.viewport().update()  # Manual UI refresh

    def calculate_dynamic_range(self, value):
        try:
            value = float(value)
            if value < 100:
                return value * 0.2
            elif 100 <= value <= 1000:
                return value * 0.1
            else:
                return value * 0.05
        except (ValueError, TypeError):
            return 0

    def on_cell_double_click(self, index):
        if not index.isValid() or self.current_view_df is None:
            return
        row = index.row()
        col = index.column()
        col_name = self.current_view_df.columns[col]
        if col_name == "Solution Label":
            return

        try:
            pivot_row = row
            current_row = 0
            for sol_label, crm_data in self._inline_crm_rows_display.items():
                pivot_idx = self.current_view_df.index[self.current_view_df['Solution Label'] == sol_label].tolist()
                if not pivot_idx:
                    continue
                pivot_idx = pivot_idx[0]
                if current_row <= row < current_row + len(crm_data) + 1:
                    if row == current_row + 1 or row == current_row + 2:
                        return
                    row = pivot_idx
                    break
                current_row += 1 + len(crm_data)

            solution_label = self.current_view_df.iloc[row]['Solution Label']
            element = col_name
            if self.use_oxide_var.isChecked():
                for el, (oxide_formula, _) in oxide_factors.items():
                    if oxide_formula == col_name:
                        element = el
                        break
            cond = (self.original_df['Solution Label'] == solution_label) & (self.original_df['Element'].str.startswith(element))
            cond &= (self.original_df['Type'] == 'Samp')
            match = self.original_df[cond]
            if match.empty:
                return

            r = match.iloc[0]
            value_column = 'Int' if self.use_int_var.isChecked() else 'Corr Con'
            value = float(r.get(value_column, 0)) / 10000
            info = [
                f"Solution: {solution_label}",
                f"Element: {col_name}",
                f"Act Wgt: {self.format_value(r.get('Act Wgt', 'N/A'))}",
                f"Act Vol: {self.format_value(r.get('Act Vol', 'N/A'))}",
                f"DF: {self.format_value(r.get('DF', 'N/A'))}",
                f"Concentration: {self.format_value(value)}"
            ]
            if element.split()[0] in oxide_factors and self.use_oxide_var.isChecked():
                formula, factor = oxide_factors[element.split()[0]]
                try:
                    oxide_value = float(value) * factor
                    info.extend([f"Oxide Formula: {formula}", f"Oxide %: {self.format_value(oxide_value)}"])
                except (ValueError, TypeError):
                    info.extend([f"Oxide Formula: {formula}", "Oxide %: N/A"])

            w = QDialog(self)
            w.setWindowTitle("Cell Information")
            w.setGeometry(200, 200, 300, 200)
            layout = QVBoxLayout(w)
            for line in info:
                layout.addWidget(QLabel(line))
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(w.accept)
            layout.addWidget(close_btn)
            w.exec()

        except Exception as e:
            self.logger.error(f"Failed to display cell info: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to display cell info: {str(e)}")

    def format_value(self, x):
        try:
            d = int(self.decimal_places.currentText())
            return f"{float(x):.{d}f}"
        except (ValueError, TypeError):
            return "" if pd.isna(x) else str(x)

    def open_row_filter_window(self):
        if self.pivot_data is None:
            QMessageBox.warning(self, "Warning", "No data to filter!")
            return
        dialog = FilterDialog(self, "Row Filter", is_row_filter=True)
        dialog.exec()

    def open_column_filter_window(self):
        if self.pivot_data is None:
            QMessageBox.warning(self, "Warning", "No data to filter!")
            return
        dialog = FilterDialog(self, "Column Filter", is_row_filter=False)
        dialog.exec()

    def reset_cache(self):
        self.pivot_data = None
        self.solution_label_order = None
        self.element_order = None
        self.column_widths.clear()
        self.cached_formatted.clear()
        self.original_df = None
        self._inline_crm_rows.clear()
        self._inline_crm_rows_display.clear()
        self.row_filter_values.clear()
        self.column_filter_values.clear()

    def clear_inline_crm(self):
        self._inline_crm_rows.clear()
        self._inline_crm_rows_display.clear()
        self.included_crms.clear()
        self.update_pivot_display()

    def show_element_plot(self):
        selected_element = self.element_selector.currentText()
        if not selected_element:
            return
        if self.current_plot_dialog:
            self.current_plot_dialog.close()
        annotations = []
        self.current_plot_dialog = PivotPlotDialog(self, selected_element, annotations)
        self.current_plot_dialog.show()

    def correct_pivot_crm(self):
        if self.pivot_data is None or self.pivot_data.empty:
            QMessageBox.warning(self, "Warning", "No pivot data available!")
            return

        selected_element = self.element_selector.currentText()
        if not selected_element:
            QMessageBox.warning(self, "Warning", "Please select an element!")
            return

        try:
            max_corr = float(self.max_correction_percent.text()) / 100
        except ValueError:
            QMessageBox.warning(self, "Warning", "Invalid max correction percent!")
            return

        try:
            self.pivot_data[selected_element] = pd.to_numeric(self.pivot_data[selected_element], errors='coerce')
            ratios = []
            problematic_labels = []

            for sol_label, crm_rows in self._inline_crm_rows_display.items():
                if sol_label not in self.included_crms or not self.included_crms[sol_label].isChecked():
                    continue
                pivot_row = self.pivot_data[self.pivot_data['Solution Label'] == sol_label]
                if pivot_row.empty:
                    continue
                pivot_val = pivot_row.iloc[0][selected_element]
                for row_data, _ in crm_rows:
                    if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                        val = row_data[self.pivot_data.columns.get_loc(selected_element)] if selected_element in self.pivot_data.columns else ""
                        if val and val.strip() and pd.notna(pivot_val):
                            try:
                                check_val = float(val)
                                pivot_val_float = float(pivot_val)
                                range_val = self.calculate_dynamic_range(check_val)
                                lower = check_val - range_val
                                upper = check_val + range_val
                                if pivot_val_float < lower:
                                    ratio = lower / pivot_val_float if pivot_val_float != 0 else float('inf')
                                    correction = abs(ratio - 1)
                                    if correction <= max_corr:
                                        ratios.append(ratio)
                                        self.pivot_data.loc[self.pivot_data['Solution Label'] == sol_label, selected_element] = pivot_val_float * ratio
                                    else:
                                        problematic_labels.append(sol_label)
                                elif pivot_val_float > upper:
                                    ratio = upper / pivot_val_float if pivot_val_float != 0 else float('inf')
                                    correction = abs(ratio - 1)
                                    if correction <= max_corr:
                                        ratios.append(ratio)
                                        self.pivot_data.loc[self.pivot_data['Solution Label'] == sol_label, selected_element] = pivot_val_float * ratio
                                    else:
                                        problematic_labels.append(sol_label)
                            except ValueError:
                                continue

            if problematic_labels:
                QMessageBox.warning(self, "Warning", f"CRM has problem in these points (correction exceeds max):\n" + "\n".join(problematic_labels))

            if not ratios:
                QMessageBox.warning(self, "Warning", "No valid ratios for correction!")
                return

            avg_ratio = np.mean(ratios)
            self.pivot_data[selected_element] = self.pivot_data[selected_element].apply(
                lambda x: x * avg_ratio if pd.notna(x) else x
            )
            self.update_pivot_display()
            if self.current_plot_dialog and self.current_plot_dialog.isVisible():
                self.current_plot_dialog.update_plot()
            QMessageBox.information(self, "Success", f"Pivot CRM corrected for {selected_element} with average ratio={avg_ratio:.3f}")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to correct Pivot CRM: {str(e)}")