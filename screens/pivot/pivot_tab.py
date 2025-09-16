import os
import platform
import re
import pandas as pd
import numpy as np
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QTableView, QAbstractItemView, 
                             QDialog, QCheckBox, QLineEdit, QFrame, QTreeView, QFileDialog, QMessageBox, QHeaderView)
from PyQt6.QtCore import Qt, QSortFilterProxyModel, QAbstractTableModel, QModelIndex
from PyQt6.QtGui import QFont, QColor, QStandardItemModel, QStandardItem, QPainter, QPen
from PyQt6.QtCharts import QChart, QChartView, QLineSeries, QScatterSeries, QValueAxis, QCategoryAxis
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font as OpenPyXLFont, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import logging

logger = logging.getLogger(__name__)

class FreezeTableWidget(QTableView):
    """A QTableView with a frozen first column that does not scroll horizontally."""
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.frozenTableView = QTableView(self)
        self.setModel(model)
        self.frozenTableView.setModel(model)
        self.init()

        self.horizontalHeader().sectionResized.connect(self.updateSectionWidth)
        self.verticalHeader().sectionResized.connect(self.updateSectionHeight)
        self.frozenTableView.verticalScrollBar().valueChanged.connect(self.frozenVerticalScroll)
        self.verticalScrollBar().valueChanged.connect(self.mainVerticalScroll)

    def init(self):
        self.frozenTableView.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.frozenTableView.verticalHeader().hide()
        self.frozenTableView.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.viewport().stackUnder(self.frozenTableView)
        
        self.frozenTableView.setStyleSheet("""
            QTableView { 
                border: none;
                selection-background-color: #999;
            }
        """)
        self.frozenTableView.setSelectionModel(self.selectionModel())
        
        for col in range(self.model().columnCount()):
            self.frozenTableView.setColumnHidden(col, col != 0)
        
        self.frozenTableView.setColumnWidth(0, self.columnWidth(0))
        self.frozenTableView.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.frozenTableView.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.frozenTableView.show()
        self.frozenTableView.viewport().repaint()
        
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        self.frozenTableView.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        
        self.updateFrozenTableGeometry()

    def updateSectionWidth(self, logicalIndex, oldSize, newSize):
        if logicalIndex == 0:
            self.frozenTableView.setColumnWidth(0, newSize)
            self.updateFrozenTableGeometry()
            self.frozenTableView.viewport().repaint()

    def updateSectionHeight(self, logicalIndex, oldSize, newSize):
        self.frozenTableView.setRowHeight(logicalIndex, newSize)
        self.frozenTableView.viewport().repaint()

    def frozenVerticalScroll(self, value):
        self.viewport().stackUnder(self.frozenTableView)
        self.verticalScrollBar().setValue(value)
        self.frozenTableView.viewport().repaint()
        self.viewport().update()

    def mainVerticalScroll(self, value):
        self.viewport().stackUnder(self.frozenTableView)
        self.frozenTableView.verticalScrollBar().setValue(value)
        self.frozenTableView.viewport().repaint()
        self.viewport().update()

    def updateFrozenTableGeometry(self):
        self.frozenTableView.setGeometry(
            self.verticalHeader().width() + self.frameWidth(),
            self.frameWidth(),
            self.columnWidth(0),
            self.viewport().height() + self.horizontalHeader().height()
        )
        self.frozenTableView.viewport().repaint()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.updateFrozenTableGeometry()
        self.frozenTableView.viewport().repaint()

    def moveCursor(self, cursorAction, modifiers):
        current = super().moveCursor(cursorAction, modifiers)
        if cursorAction == QAbstractItemView.CursorAction.MoveLeft and current.column() > 0:
            visual_x = self.visualRect(current).topLeft().x()
            if visual_x < self.frozenTableView.columnWidth(0):
                new_value = self.horizontalScrollBar().value() + visual_x - self.frozenTableView.columnWidth(0)
                self.horizontalScrollBar().setValue(int(new_value))
        self.frozenTableView.viewport().repaint()
        return current

    def scrollTo(self, index, hint=QAbstractItemView.ScrollHint.EnsureVisible):
        if index.column() > 0:
            super().scrollTo(index, hint)
        self.frozenTableView.viewport().repaint()

class PivotTableModel(QAbstractTableModel):
    """Custom table model for pivot table, optimized for large datasets."""
    def __init__(self, pivot_tab, df=None, crm_rows=None):
        super().__init__()
        self.pivot_tab = pivot_tab
        self._df = df if df is not None else pd.DataFrame()
        self._crm_rows = crm_rows if crm_rows is not None else []
        self._row_info = []
        self._column_widths = {}
        self._build_row_info()

    def set_data(self, df, crm_rows=None):
        self.beginResetModel()
        self._df = df.copy()
        self._crm_rows = crm_rows if crm_rows is not None else []
        self._build_row_info()
        self.endResetModel()

    def _build_row_info(self):
        self._row_info = []
        for row_idx in range(len(self._df)):
            self._row_info.append({'type': 'pivot', 'index': row_idx})
            sol_label = self._df.iloc[row_idx]['Solution Label']
            for grp_idx, (sl, cdata) in enumerate(self._crm_rows):
                if sl == sol_label:
                    for sub in range(len(cdata)):
                        self._row_info.append({'type': 'crm', 'group': grp_idx, 'sub': sub})
                    break

    def rowCount(self, parent=QModelIndex()):
        return len(self._row_info)

    def columnCount(self, parent=QModelIndex()):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or index.row() >= len(self._row_info):
            return None

        row = index.row()
        col = index.column()
        col_name = self._df.columns[col]
        info = self._row_info[row]

        is_crm_row = False
        is_diff_row = False
        crm_row_data = None
        tags = None
        pivot_row = row

        if info['type'] == 'pivot':
            pivot_row = info['index']
        else:
            grp = info['group']
            sub = info['sub']
            _, crm_data = self._crm_rows[grp]
            if sub == 0:
                is_crm_row = True
                crm_row_data = crm_data[0][0]
                tags = crm_data[0][1]
            elif sub == 1:
                is_diff_row = True
                crm_row_data = crm_data[1][0]
                tags = crm_data[1][1]
            pivot_row = self._df.index[self._df['Solution Label'] == self._crm_rows[grp][0]].tolist()[0]

        if role == Qt.ItemDataRole.DisplayRole:
            dec = int(self.pivot_tab.decimal_places.currentText())
            if is_crm_row or is_diff_row:
                value = crm_row_data[col]
                return str(value) if value else ""
            else:
                value = self._df.iloc[pivot_row, col]
                if col_name != "Solution Label" and pd.notna(value):
                    try:
                        if self.pivot_tab.use_oxide_var.isChecked() and col_name in self.pivot_tab.oxide_factors:
                            _, factor = self.pivot_tab.oxide_factors[col_name]
                            value = float(value) * factor
                        return f"{float(value):.{dec}f}"
                    except (ValueError, TypeError):
                        return "" if pd.isna(value) else str(value)
                return str(value) if pd.notna(value) else ""

        elif role == Qt.ItemDataRole.BackgroundRole:
            if is_crm_row:
                return QColor("#FFF5E4")
            elif is_diff_row and tags:
                if tags[col] == "in_range":
                    return QColor("#ECFFC4")
                elif tags[col] == "out_range":
                    return QColor("#FFCCCC")
                return QColor("#E6E6FA")
            return QColor("#f9f9f9") if pivot_row % 2 == 0 else QColor("white")

        elif role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignLeft if col_name == "Solution Label" else Qt.AlignmentFlag.AlignCenter

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._df.columns[section])
            return str(section + 1)
        return None

    def set_column_width(self, col, width):
        self._column_widths[col] = width

class FilterDialog(QDialog):
    """Dialog for row/column filtering."""
    def __init__(self, parent, title, is_row_filter=True):
        super().__init__(parent)
        self.parent = parent
        self.is_row_filter = is_row_filter
        self.setWindowTitle(title)
        self.setGeometry(200, 200, 300, 460)
        self.setModal(True)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select Field:"))
        
        self.field_combo = QComboBox()
        if self.is_row_filter:
            self.field_combo.addItems(["Solution Label", "Element"])
        else:
            self.field_combo.addItems(["Element"])
        layout.addWidget(self.field_combo)

        self.tree_view = QTreeView()
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["Value", "Select"])
        self.tree_view.setModel(self.model)
        self.tree_view.setRootIsDecorated(False)
        self.tree_view.header().setStretchLastSection(False)
        self.tree_view.header().resizeSection(0, 160)
        self.tree_view.header().resizeSection(1, 80)
        layout.addWidget(self.tree_view)

        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(lambda: self.set_all_checks(True))
        button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(lambda: self.set_all_checks(False))
        button_layout.addWidget(deselect_all_btn)
        
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.apply_filters)
        button_layout.addWidget(apply_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        self.field_combo.currentTextChanged.connect(self.update_tree)
        self.update_tree()

    def update_tree(self):
        self.model.clear()
        self.model.setHorizontalHeaderLabels(["Value", "Select"])
        field = self.field_combo.currentText()
        if not field:
            return
        
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        uniques = (self.parent.solution_label_order if self.is_row_filter and field == "Solution Label" else
                  sorted(self.parent.pivot_data[field].astype(str).unique()) if field in self.parent.pivot_data.columns else
                  self.parent.element_order)
        
        if field not in var_dict:
            var_dict[field] = {v: True for v in uniques}
        
        for val in uniques:
            value_item = QStandardItem(str(val))
            check_item = QStandardItem()
            check_item.setCheckable(True)
            check_item.setCheckState(Qt.CheckState.Checked if var_dict[field].get(val, True) else Qt.CheckState.Unchecked)
            self.model.appendRow([value_item, check_item])
        
        self.tree_view.clicked.connect(self.toggle_check)

    def toggle_check(self, index):
        if index.column() != 1:
            return
        val = self.model.item(index.row(), 0).text()
        field = self.field_combo.currentText()
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        var_dict[field][val] = not var_dict[field].get(val, True)
        self.model.item(index.row(), 1).setCheckState(
            Qt.CheckState.Checked if var_dict[field][val] else Qt.CheckState.Unchecked)

    def set_all_checks(self, value):
        field = self.field_combo.currentText()
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        if field not in var_dict:
            return
        for val in var_dict[field]:
            var_dict[field][val] = value
        self.update_tree()

    def apply_filters(self):
        self.parent.update_pivot_display()
        self.accept()

class PivotPlotDialog(QDialog):
    """Dialog for plotting CRM data with a separate difference plot using PyQt6 QtCharts."""
    def __init__(self, parent, selected_element):
        super().__init__(parent)
        self.parent = parent
        self.selected_element = selected_element
        self.setWindowTitle(f"CRM Plot for {selected_element}")
        self.setGeometry(200, 200, 1000, 800)
        self.setModal(False)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        control_frame = QHBoxLayout()
        self.show_check_crm = QCheckBox("Show Check CRM", checked=True)
        self.show_check_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_check_crm)
        
        self.show_pivot_crm = QCheckBox("Show Pivot CRM", checked=True)
        self.show_pivot_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_pivot_crm)
        
        self.show_middle = QCheckBox("Show Middle", checked=True)
        self.show_middle.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_middle)
        
        self.show_range = QCheckBox("Show Range", checked=True)
        self.show_range.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_range)
        
        self.show_diff = QCheckBox("Show Difference", checked=True)
        self.show_diff.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_diff)
        
        control_frame.addWidget(QLabel("Max Correction (%):"))
        self.max_correction_percent = QLineEdit("32")
        control_frame.addWidget(self.max_correction_percent)
        
        correct_btn = QPushButton("Correct Pivot CRM")
        correct_btn.clicked.connect(self.correct_crm_callback)
        control_frame.addWidget(correct_btn)
        
        select_crms_btn = QPushButton("Select CRMs")
        select_crms_btn.clicked.connect(self.open_select_crms_window)
        control_frame.addWidget(select_crms_btn)
        
        layout.addLayout(control_frame)
        
        self.chart = QChart()
        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        layout.addWidget(self.chart_view)
        
        self.diff_chart = QChart()
        self.diff_chart_view = QChartView(self.diff_chart)
        self.diff_chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        layout.addWidget(self.diff_chart_view)
        
        self.update_plot()

    def correct_crm_callback(self):
        self.parent.correct_pivot_crm()
        self.update_plot()

    def update_plot(self):
        if not self.selected_element or self.selected_element not in self.parent.pivot_data.columns:
            QMessageBox.warning(self, "Warning", f"Element '{self.selected_element}' not found in pivot data!")
            return

        try:
            self.chart.removeAllSeries()
            self.diff_chart.removeAllSeries()
            for axis in self.chart.axes() + self.diff_chart.axes():
                self.chart.removeAxis(axis)
                self.diff_chart.removeAxis(axis)

            valid_pairs = []
            for sol_label, crm_rows in self.parent._inline_crm_rows_display.items():
                if sol_label not in self.parent.included_crms or not self.parent.included_crms[sol_label].isChecked():
                    continue
                pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                if pivot_row.empty:
                    continue
                pivot_val = pivot_row.iloc[0][self.selected_element]
                for row_data, _ in crm_rows:
                    if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                        val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                        if val and val.strip():
                            try:
                                crm_val = float(val)
                                if pd.notna(pivot_val):
                                    pivot_val_float = float(pivot_val)
                                    valid_pairs.append((sol_label, crm_val, pivot_val_float))
                            except ValueError:
                                continue

            valid_pairs.sort(key=lambda x: x[0])
            crm_display_values = [pair[1] for pair in valid_pairs]
            crm_pivot_values = [pair[2] for pair in valid_pairs]
            labels = [pair[0] for pair in valid_pairs]

            if not valid_pairs:
                pivot_data = self.parent.pivot_data[self.parent.pivot_data['Solution Label'].str.contains('CRM|par', case=False, na=False)]
                if not pivot_data.empty and self.show_pivot_crm.isChecked():
                    series = QLineSeries()
                    series.setName(f"{self.selected_element} (Pivot CRM)")
                    series.setPen(QPen(QColor("blue"), 2))
                    for i, (_, row) in enumerate(pivot_data.iterrows()):
                        val = row[self.selected_element]
                        try:
                            if pd.notna(val):
                                series.append(i, float(val))
                        except ValueError:
                            continue
                    if series.count() > 0:
                        self.chart.addSeries(series)
                        axis_x = QCategoryAxis()
                        for i, (_, row) in enumerate(pivot_data.iterrows()):
                            axis_x.append(row['Solution Label'], i)
                        axis_x.setLabelsAngle(45)
                        axis_y = QValueAxis()
                        axis_y.setTitleText(f"{self.selected_element} Value")
                        self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
                        self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
                        series.attachAxis(axis_x)
                        series.attachAxis(axis_y)
                        self.chart.setTitle(f"Pivot CRM Values for {self.selected_element}")
                        self.chart_view.update()
                        self.diff_chart.setTitle("No difference data available")
                        self.diff_chart_view.update()
                        return
                self.chart.setTitle(f"No valid CRM data for {self.selected_element}")
                self.chart_view.update()
                self.diff_chart.setTitle("No difference data available")
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"No valid CRM data for {self.selected_element}")
                return

            if not all(isinstance(v, (int, float)) and not np.isnan(v) for v in crm_display_values + crm_pivot_values):
                self.chart.setTitle(f"Invalid data for {self.selected_element}")
                self.chart_view.update()
                self.diff_chart.setTitle("Invalid difference data")
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"Invalid or non-numeric data for {self.selected_element}")
                return

            middle_values = [(disp_val + piv_val) / 2 if disp_val != 0 and piv_val != 0 else disp_val or piv_val or 0
                             for disp_val, piv_val in zip(crm_display_values, crm_pivot_values)]
            range_lower = [val - self.parent.calculate_dynamic_range(val) for val in crm_display_values]
            range_upper = [val + self.parent.calculate_dynamic_range(val) for val in crm_display_values]
            diff_values = [((crm_val - piv_val) / crm_val) * 100 if crm_val != 0 else 0
                           for crm_val, piv_val in zip(crm_display_values, crm_pivot_values)]

            axis_x = QCategoryAxis()
            for i, label in enumerate(labels):
                axis_x.append(label, i)
            axis_x.setLabelsAngle(45)
            self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)

            axis_y = QValueAxis()
            axis_y.setTitleText(f"{self.selected_element} Value")
            self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)

            axis_x_diff = QCategoryAxis()
            for i, label in enumerate(labels):
                axis_x_diff.append(label, i)
            axis_x_diff.setLabelsAngle(45)
            self.diff_chart.addAxis(axis_x_diff, Qt.AlignmentFlag.AlignBottom)

            axis_y_diff = QValueAxis()
            axis_y_diff.setTitleText("Difference (%)")
            axis_y_diff.setLabelFormat("%.1f")
            self.diff_chart.addAxis(axis_y_diff, Qt.AlignmentFlag.AlignLeft)

            plotted = False
            if self.show_check_crm.isChecked() and any(v != 0 for v in crm_display_values):
                series = QLineSeries()
                series.setName(f"{self.selected_element} (Check CRM)")
                series.setPen(QPen(QColor("red"), 2, Qt.PenStyle.DashLine))
                for i, val in enumerate(crm_display_values):
                    series.append(i, val)
                self.chart.addSeries(series)
                series.attachAxis(axis_x)
                series.attachAxis(axis_y)
                if self.show_range.isChecked():
                    range_series_lower = QLineSeries()
                    range_series_lower.setName("Check CRM Range Lower")
                    range_series_lower.setPen(QPen(QColor("red"), 1, Qt.PenStyle.DotLine))
                    range_series_upper = QLineSeries()
                    range_series_upper.setName("Check CRM Range Upper")
                    range_series_upper.setPen(QPen(QColor("red"), 1, Qt.PenStyle.DotLine))
                    for i, lower in enumerate(range_lower):
                        range_series_lower.append(i, lower)
                    for i, upper in enumerate(range_upper):
                        range_series_upper.append(i, upper)
                    self.chart.addSeries(range_series_lower)
                    self.chart.addSeries(range_series_upper)
                    range_series_lower.attachAxis(axis_x)
                    range_series_lower.attachAxis(axis_y)
                    range_series_upper.attachAxis(axis_x)
                    range_series_upper.attachAxis(axis_y)
                plotted = True

            if self.show_pivot_crm.isChecked() and any(v != 0 for v in crm_pivot_values):
                series = QLineSeries()
                series.setName(f"{self.selected_element} (Pivot CRM)")
                series.setPen(QPen(QColor("blue"), 2))
                scatter = QScatterSeries()
                scatter.setName("")
                scatter.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter.setMarkerSize(8)
                scatter.setPen(QPen(QColor("blue")))
                for i, val in enumerate(crm_pivot_values):
                    series.append(i, val)
                    scatter.append(i, val)
                self.chart.addSeries(series)
                self.chart.addSeries(scatter)
                series.attachAxis(axis_x)
                series.attachAxis(axis_y)
                scatter.attachAxis(axis_x)
                scatter.attachAxis(axis_y)
                plotted = True

            if self.show_middle.isChecked() and any(v != 0 for v in middle_values):
                series = QLineSeries()
                series.setName(f"{self.selected_element} (Middle)")
                series.setPen(QPen(QColor("green"), 2, Qt.PenStyle.DotLine))
                for i, val in enumerate(middle_values):
                    series.append(i, val)
                self.chart.addSeries(series)
                series.attachAxis(axis_x)
                series.attachAxis(axis_y)
                plotted = True

            if self.show_diff.isChecked() and any(v != 0 for v in diff_values):
                series = QLineSeries()
                series.setName("Difference (%)")
                series.setPen(QPen(QColor("purple"), 2))
                scatter = QScatterSeries()
                scatter.setName("")
                scatter.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter.setMarkerSize(8)
                scatter.setPen(QPen(QColor("purple")))
                for i, val in enumerate(diff_values):
                    series.append(i, val)
                    scatter.append(i, val)
                self.diff_chart.addSeries(series)
                self.diff_chart.addSeries(scatter)
                series.attachAxis(axis_x_diff)
                series.attachAxis(axis_y_diff)
                scatter.attachAxis(axis_x_diff)
                scatter.attachAxis(axis_y_diff)
                plotted = True

            if not plotted:
                self.chart.setTitle(f"No data plotted for {self.selected_element}")
                self.diff_chart.setTitle("No difference data plotted")
                self.chart_view.update()
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"No data could be plotted for {self.selected_element}")
                return

            self.chart.setTitle(f"CRM Values for {self.selected_element}")
            self.diff_chart.setTitle(f"Difference (%) for {self.selected_element}")
            self.chart_view.update()
            self.diff_chart_view.update()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to update plot: {str(e)}")

    def open_select_crms_window(self):
        w = QDialog(self)
        w.setWindowTitle("Select CRMs to Include")
        w.setGeometry(200, 200, 300, 400)
        w.setModal(True)
        layout = QVBoxLayout(w)

        tree_view = QTreeView()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Label", "Include"])
        tree_view.setModel(model)
        tree_view.setRootIsDecorated(False)
        tree_view.header().resizeSection(0, 160)
        tree_view.header().resizeSection(1, 80)

        for label in sorted(self.parent.included_crms.keys()):
            value_item = QStandardItem(label)
            check_item = QStandardItem()
            check_item.setCheckable(True)
            check_item.setCheckState(Qt.CheckState.Checked if self.parent.included_crms[label].isChecked() else Qt.CheckState.Unchecked)
            model.appendRow([value_item, check_item])

        tree_view.clicked.connect(lambda index: self.toggle_crm_check(index, model))
        layout.addWidget(tree_view)

        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(lambda: self.set_all_crms(True, model))
        button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(lambda: self.set_all_crms(False, model))
        button_layout.addWidget(deselect_all_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(w.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        w.exec()

    def toggle_crm_check(self, index, model):
        if index.column() != 1:
            return
        label = model.item(index.row(), 0).text()
        if label in self.parent.included_crms:
            self.parent.included_crms[label].setChecked(not self.parent.included_crms[label].isChecked())
            model.item(index.row(), 1).setCheckState(
                Qt.CheckState.Checked if self.parent.included_crms[label].isChecked() else Qt.CheckState.Unchecked)
            self.update_plot()

    def set_all_crms(self, value, model):
        for label, checkbox in self.parent.included_crms.items():
            checkbox.setChecked(value)
        model.clear()
        model.setHorizontalHeaderLabels(["Label", "Include"])
        for label in sorted(self.parent.included_crms.keys()):
            value_item = QStandardItem(label)
            check_item = QStandardItem()
            check_item.setCheckable(True)
            check_item.setCheckState(Qt.CheckState.Checked if self.parent.included_crms[label].isChecked() else Qt.CheckState.Unchecked)
            model.appendRow([value_item, check_item])
        self.update_plot()

class PivotTab(QWidget):
    """PivotTab with inline CRM rows, difference coloring, and plot visualization."""
    oxide_factors = {
        'Si': ('SiO2', 60.0843 / 28.0855),
        'Ti': ('TiO2', 79.8658 / 47.867),
        'Al': ('Al2O3', 101.9613 / (2 * 26.9815)),
        'Fe': ('Fe2O3', 159.6872 / (2 * 55.845)),
        'Mn': ('MnO', 70.9374 / 54.9380),
        'Mg': ('MgO', 40.3044 / 24.305),
        'Ca': ('CaO', 56.0774 / 40.078),
        'Na': ('Na2O', 61.9789 / (2 * 22.9898)),
        'K': ('K2O', 94.1960 / (2 * 39.0983)),
        'P': ('P2O5', 141.9445 / (2 * 30.9738)),
        'Cr': ('Cr2O3', 151.9902 / (2 * 51.9961)),
        'Ni': ('NiO', 74.6928 / 58.6934),
        'Cu': ('CuO', 79.5454 / 63.546),
        'Zn': ('ZnO', 81.3794 / 65.38),
        'Ba': ('BaO', 153.3294 / 137.327),
        'Sr': ('SrO', 103.6194 / 87.62),
        'V': ('V2O5', 181.8802 / (2 * 50.9415)),
        'Zr': ('ZrO2', 123.2182 / 91.224),
        'Nb': ('Nb2O5', 265.8098 / (2 * 92.9064)),
        'Y': ('Y2O3', 225.8102 / (2 * 88.9058)),
        'La': ('La2O3', 325.8092 / (2 * 138.9055)),
        'Ce': ('CeO2', 172.1148 / 140.116),
        'Nd': ('Nd2O3', 336.482 / (2 * 144.242)),
        'Pb': ('PbO', 223.1992 / 207.2),
        'Th': ('ThO2', 248.0722 / 232.038),
        'U': ('U3O8', 842.088 / (3 * 238.029)),
    }

    def __init__(self, app, parent_frame):
        super().__init__(parent_frame)
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
        
        self.use_int_var.toggled.connect(self.create_pivot)
        control_layout.addWidget(self.use_int_var)
        
        self.use_oxide_var.toggled.connect(self.create_pivot)
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
        check_rm_btn.clicked.connect(self.check_rm)
        control_layout.addWidget(check_rm_btn)
        
        clear_crm_btn = QPushButton("Clear CRM")
        clear_crm_btn.clicked.connect(self.clear_inline_crm)
        control_layout.addWidget(clear_crm_btn)
        
        export_btn = QPushButton("Export")
        export_btn.clicked.connect(self.export_pivot)
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

    def create_pivot(self):
        df = self.app.get_data()
        if df is None or df.empty:
            QMessageBox.warning(self, "Warning", "No data to display!")
            return

        try:
            self.original_df = df.copy()
            df_filtered = df[df['Type'].isin(['Samp', 'Sample'])].copy()
            df_filtered['original_index'] = df_filtered.index
            grp = (df_filtered['Element'] != df_filtered['Element'].shift()).cumsum()
            df_filtered['Element'] = df_filtered['Element'].str.split('_').str[0]
            df_filtered['unique_id'] = df_filtered.groupby(['Solution Label', 'Element']).cumcount()

            self.solution_label_order = df_filtered['Solution Label'].drop_duplicates().tolist()
            self.element_order = df_filtered['Element'].drop_duplicates().tolist()
            self.element_selector.clear()
            self.element_selector.addItems([""] + self.element_order)

            value_column = 'Int' if self.use_int_var.isChecked() else 'Corr Con'
            if value_column not in df_filtered.columns:
                QMessageBox.warning(self, "Error", f"Column '{value_column}' not found in data!")
                return

            pivot_df = df_filtered.pivot_table(
                index=['Solution Label', 'unique_id'],
                columns='Element',
                values=value_column,
                aggfunc='first'
            ).reset_index()
            pivot_df = pivot_df.merge(
                df_filtered[['original_index', 'Solution Label', 'unique_id']],
                on=['Solution Label', 'unique_id'],
                how='left'
            ).sort_values('original_index').drop(columns=['original_index', 'unique_id']).drop_duplicates()
            
            if self.use_oxide_var.isChecked():
                for col in pivot_df.columns:
                    if col != 'Solution Label' and col in self.oxide_factors:
                        _, factor = self.oxide_factors[col]
                        pivot_df[col] = pd.to_numeric(pivot_df[col], errors='coerce') * factor
            
            self.pivot_data = pivot_df
            self.column_widths.clear()
            self.cached_formatted.clear()
            self._inline_crm_rows.clear()
            self._inline_crm_rows_display.clear()
            self.row_filter_values.clear()
            self.column_filter_values.clear()
            self.update_pivot_display()

        except Exception as e:
            QMessageBox.warning(self, "Pivot Error", f"Failed to create pivot table: {str(e)}")

    def update_pivot_display(self):
        if self.pivot_data is None or self.pivot_data.empty:
            self.status_label.setText("No data loaded")
            self.table_view.setModel(None)
            self.table_view.frozenTableView.setModel(None)
            return

        df = self.pivot_data.copy()
        s = self.search_var.text().strip().lower()
        if s:
            mask = df.apply(lambda r: r.astype(str).str.lower().str.contains(s, na=False).any(), axis=1)
            df = df[mask]

        for field, values in self.row_filter_values.items():
            if field in df.columns:
                selected = [k for k, v in values.items() if v]
                if selected:
                    df = df[df[field].isin(selected)]

        selected_cols = ['Solution Label']
        for field, values in self.column_filter_values.items():
            if field == 'Element':
                selected_cols.extend([k for k, v in values.items() if v and k in df.columns])
        if len(selected_cols) > 1:
            df = df[selected_cols]

        df = df.reset_index(drop=True)
        self.current_view_df = df
        self._inline_crm_rows_display = self._build_crm_row_lists_for_columns(list(df.columns))

        crm_rows = []
        for sol_label in df['Solution Label']:
            if sol_label in self._inline_crm_rows_display:
                crm_rows.append((sol_label, self._inline_crm_rows_display[sol_label]))

        model = PivotTableModel(self, df, crm_rows)
        self.table_view.setModel(model)
        self.table_view.frozenTableView.setModel(model)
        model.layoutChanged.emit()
        for col, width in self.column_widths.items():
            self.table_view.horizontalHeader().resizeSection(col, width)
        self.status_label.setText("Data loaded successfully")

    def _build_crm_row_lists_for_columns(self, columns):
        crm_display = {}
        dec = int(self.decimal_places.currentText())
        try:
            min_diff = float(self.diff_min.text())
            max_diff = float(self.diff_max.text())
        except ValueError:
            min_diff, max_diff = -12, 12

        has_act_vol = 'Act Vol' in self.original_df.columns if self.original_df is not None else False
        has_act_wg = 'Act Wgt' in self.original_df.columns if self.original_df is not None else False
        has_coeff_1 = 'Coeff 1' in self.original_df.columns if self.original_df is not None else False
        has_coeff_2 = 'Coeff 2' in self.original_df.columns if self.original_df is not None else False

        for sol_label, list_of_dicts in self._inline_crm_rows.items():
            crm_display[sol_label] = []
            pivot_row = self.pivot_data[self.pivot_data['Solution Label'] == sol_label]
            if pivot_row.empty:
                continue
            pivot_values = pivot_row.iloc[0].to_dict()

            element_params = {}
            if self.original_df is not None:
                sample_rows = self.original_df[
                    (self.original_df['Solution Label'] == sol_label) & 
                    (self.original_df['Type'].isin(['Sample', 'Samp']))
                ]
                for _, row in sample_rows.iterrows():
                    element = row['Element'].split('_')[0]
                    element_params[element] = {
                        'Act Vol': row['Act Vol'] if has_act_vol else 1.0,
                        'Act Wgt': row['Act Wgt'] if has_act_wg else 1.0,
                        'Coeff 1': row['Coeff 1'] if has_coeff_1 else 0.0,
                        'Coeff 2': row['Coeff 2'] if has_coeff_2 else 1.0
                    }

            for d in list_of_dicts:
                crm_row_list = []
                for col in columns:
                    if col == 'Solution Label':
                        crm_row_list.append(f"{d.get('Solution Label', sol_label)} CRM")
                    else:
                        val = d.get(col, "")
                        if pd.isna(val) or val == "":
                            crm_row_list.append("")
                        else:
                            try:
                                element_symbol = col.split('_')[0]
                                params = element_params.get(element_symbol, {
                                    'Act Vol': 1.0, 'Act Wgt': 1.0, 'Coeff 1': 0.0, 'Coeff 2': 1.0
                                })
                                act_vol = params['Act Vol']
                                act_wg = params['Act Wgt']
                                coeff_1 = params['Coeff 1']
                                coeff_2 = params['Coeff 2']
                                if self.use_int_var.isChecked() and act_vol != 0 and act_wg != 0:
                                    int_value = coeff_2 * (float(val) / (act_vol / act_wg)) + coeff_1
                                    if element_symbol in self.oxide_factors and self.use_oxide_var.isChecked():
                                        _, factor = self.oxide_factors[element_symbol]
                                        oxide_int_value = int_value * factor
                                        crm_row_list.append(f"{oxide_int_value:.{dec}f}")
                                    else:
                                        crm_row_list.append(f"{int_value:.{dec}f}")
                                else:
                                    if element_symbol in self.oxide_factors and self.use_oxide_var.isChecked():
                                        _, factor = self.oxide_factors[element_symbol]
                                        oxide_val = float(val) * factor
                                        crm_row_list.append(f"{oxide_val:.{dec}f}")
                                    else:
                                        crm_row_list.append(f"{float(val):.{dec}f}")
                            except Exception:
                                crm_row_list.append(str(val))
                crm_display[sol_label].append((crm_row_list, ["crm"] * len(columns)))

                diff_row_list = []
                diff_tags = []
                for col in columns:
                    if col == 'Solution Label':
                        diff_row_list.append(f"{sol_label} Diff (%)")
                        diff_tags.append("diff")
                    else:
                        pivot_val = pivot_values.get(col, None)
                        crm_val = d.get(col, None)
                        if pivot_val is not None and crm_val is not None:
                            try:
                                element_symbol = col.split('_')[0]
                                params = element_params.get(element_symbol, {
                                    'Act Vol': 1.0, 'Act Wgt': 1.0, 'Coeff 1': 0.0, 'Coeff 2': 1.0
                                })
                                act_vol = params['Act Vol']
                                act_wg = params['Act Wgt']
                                coeff_1 = params['Coeff 1']
                                coeff_2 = params['Coeff 2']
                                pivot_val = float(pivot_val)
                                if element_symbol in self.oxide_factors and self.use_oxide_var.isChecked():
                                    _, pivot_factor = self.oxide_factors[element_symbol]
                                    pivot_val *= pivot_factor
                                if self.use_int_var.isChecked() and act_vol != 0 and act_wg != 0:
                                    crm_val = coeff_2 * (float(crm_val) / (act_vol / act_wg)) + coeff_1
                                if element_symbol in self.oxide_factors and self.use_oxide_var.isChecked():
                                    _, crm_factor = self.oxide_factors[element_symbol]
                                    crm_val = float(crm_val) * crm_factor
                                else:
                                    crm_val = float(crm_val)
                                if crm_val != 0:
                                    diff = ((crm_val - pivot_val) / crm_val) * 100
                                    diff_row_list.append(f"{diff:.{dec}f}")
                                    diff_tags.append("in_range" if min_diff <= diff <= max_diff else "out_range")
                                else:
                                    diff_row_list.append("N/A")
                                    diff_tags.append("diff")
                            except Exception:
                                diff_row_list.append("")
                                diff_tags.append("diff")
                        else:
                            diff_row_list.append("")
                            diff_tags.append("diff")
                crm_display[sol_label].append((diff_row_list, diff_tags))

        return crm_display

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
            element = col_name.split('_')[0]
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
            if element.split()[0] in self.oxide_factors and self.use_oxide_var.isChecked():
                formula, factor = self.oxide_factors[element.split()[0]]
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

    def export_pivot(self):
        if self.pivot_data is None or self.pivot_data.empty:
            QMessageBox.warning(self, "Warning", "No data to export!")
            return

        try:
            filtered = self.pivot_data.copy()
            for field, values in self.row_filter_values.items():
                if field in filtered.columns:
                    selected = [k for k, v in values.items() if v]
                    if selected:
                        filtered = filtered[filtered[field].isin(selected)]

            selected_cols = ['Solution Label']
            for field, values in self.column_filter_values.items():
                if field == 'Element':
                    selected_cols.extend([k for k, v in values.items() if v and k in filtered.columns])
            if len(selected_cols) > 1:
                filtered = filtered[selected_cols]

            filtered['Solution Label'] = pd.Categorical(filtered['Solution Label'], categories=self.solution_label_order, ordered=True)
            filtered = filtered.sort_values('Solution Label').reset_index(drop=True)

            file_path = QFileDialog.getSaveFileName(self, "Save Pivot Table", "", "Excel files (*.xlsx)")[0]
            if not file_path:
                return

            wb = Workbook()
            ws = wb.active
            ws.title = "Pivot Table"

            header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            first_col_fill = PatternFill(start_color="FFF5E4", end_color="FFF5E4", fill_type="solid")
            odd_fill = PatternFill(start_color="f5f5f5", end_color="f5f5f5", fill_type="solid")
            even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            diff_in_range_fill = PatternFill(start_color="ECFFC4", end_color="ECFFC4", fill_type="solid")
            diff_out_range_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            header_font = OpenPyXLFont(name="Segoe UI", size=12, bold=True)
            cell_font = OpenPyXLFont(name="Segoe UI", size=12)
            cell_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

            headers = list(filtered.columns)
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill = header_fill
                c.font = header_font
                c.alignment = cell_align
                c.border = thin_border
                ws.column_dimensions[get_column_letter(ci)].width = 15

            dec = int(self.decimal_places.currentText())
            try:
                min_diff = float(self.diff_min.text())
                max_diff = float(self.diff_max.text())
            except ValueError:
                min_diff, max_diff = -12, 12

            row_idx = 2
            for _, row in filtered.iterrows():
                sol_label = row['Solution Label']
                for ci, val in enumerate(row, 1):
                    cell = ws.cell(row=row_idx, column=ci)
                    try:
                        cell.value = round(float(val), dec)
                        cell.number_format = f"0.{''.join(['0']*dec)}" if dec > 0 else "0"
                    except (ValueError, TypeError):
                        cell.value = "" if pd.isna(val) else str(val)
                    cell.font = cell_font
                    cell.alignment = cell_align
                    cell.border = thin_border
                    cell.fill = first_col_fill if ci == 1 else (even_fill if (row_idx - 1) % 2 == 0 else odd_fill)
                row_idx += 1

                if sol_label in self._inline_crm_rows_display:
                    for crm_row_data, tags in self._inline_crm_rows_display[sol_label]:
                        for ci, val in enumerate(crm_row_data, 1):
                            cell = ws.cell(row=row_idx, column=ci)
                            cell.value = str(val) if val else ""
                            cell.font = cell_font
                            cell.alignment = cell_align
                            cell.border = thin_border
                            if crm_row_data[0].endswith("CRM"):
                                cell.fill = first_col_fill if ci == 1 else PatternFill(start_color="FFF5E4", end_color="FFF5E4", fill_type="solid")
                            elif crm_row_data[0].endswith("Diff (%)"):
                                if tags[ci-1] == "in_range":
                                    cell.fill = diff_in_range_fill
                                elif tags[ci-1] == "out_range":
                                    cell.fill = diff_out_range_fill
                                else:
                                    cell.fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
                        row_idx += 1

            wb.save(file_path)
            QMessageBox.information(self, "Success", "Pivot table exported successfully!")
            if QMessageBox.question(self, "Open File", "Open the saved Excel file?") == QMessageBox.StandardButton.Yes:
                try:
                    if platform.system() == "Windows":
                        os.startfile(file_path)
                    elif platform.system() == "Darwin":
                        os.system(f"open '{file_path}'")
                    else:
                        os.system(f"xdg-open '{file_path}'")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Failed to open file: {str(e)}")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to export: {str(e)}")

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

    def check_rm(self):
        if self.pivot_data is None or self.pivot_data.empty:
            QMessageBox.warning(self, "Warning", "No pivot data available!")
            return

        try:
            conn = self.app.crm_tab.conn
            if conn is None:
                self.app.crm_tab.init_db()
                conn = self.app.crm_tab.conn
                if conn is None:
                    QMessageBox.warning(self, "Error", "Failed to connect to CRM database!")
                    return

            crm_rows = self.pivot_data[self.pivot_data['Solution Label'].str.contains('CRM|par', case=False, na=False)].copy()
            if crm_rows.empty:
                QMessageBox.information(self, "Info", "No CRM or par rows found in pivot data!")
                return

            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(crm)")
            cols = [x[1] for x in cursor.fetchall()]
            required = {'CRM ID', 'Element', 'Sort Grade', 'Analysis Method'}
            if not required.issubset(cols):
                QMessageBox.warning(self, "Error", "CRM table missing required columns!")
                return

            dec = int(self.decimal_places.currentText())
            self._inline_crm_rows.clear()
            self.included_crms.clear()
            for _, row in crm_rows.iterrows():
                label = row['Solution Label']
                m = re.search(r'\d+', str(label))
                if not m:
                    continue
                crm_id_string = f"OREAS {m.group()}"
                analysis_method = '4-Acid Digestion' if 'CRM' in str(label).upper() else 'Aqua Regia Digestion'

                cursor.execute(
                    "SELECT [Element], [Sort Grade] FROM crm WHERE [CRM ID] = ? AND [Analysis Method] = ?",
                    (crm_id_string, analysis_method)
                )
                crm_data = cursor.fetchall()
                if not crm_data:
                    continue

                crm_dict = {}
                for element_full, sort_grade in crm_data:
                    symbol = element_full.split(',')[-1].strip() if ',' in element_full else element_full.split()[-1].strip()
                    try:
                        crm_dict[symbol] = float(sort_grade)
                    except (ValueError, TypeError):
                        pass

                crm_values = {'Solution Label': crm_id_string}
                for col in self.element_order:
                    if col == 'Solution Label' or col not in self.pivot_data.columns:
                        continue
                    element_symbol = col.split(' ')[0].split('_')[0]
                    if element_symbol in crm_dict:
                        if element_symbol in self.oxide_factors and self.use_oxide_var.isChecked():
                            _, factor = self.oxide_factors[element_symbol]
                            crm_values[col] = crm_dict[element_symbol] * factor
                        else:
                            crm_values[col] = crm_dict[element_symbol]
                if len(crm_values) > 1:
                    self._inline_crm_rows.setdefault(label, []).append(crm_values)
                    self.included_crms[label] = QCheckBox(label, checked=True)

            if not self._inline_crm_rows:
                QMessageBox.information(self, "Info", "No matching CRM elements found for comparison!")
                return

            self._inline_crm_rows_display = self._build_crm_row_lists_for_columns(list(self.pivot_data.columns))
            self.update_pivot_display()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to check RM: {str(e)}")

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

    def show_element_plot(self):
        if self.pivot_data is None or self.pivot_data.empty:
            QMessageBox.warning(self, "Warning", "No data to plot!")
            return

        selected_element = self.element_selector.currentText()
        if not selected_element or selected_element not in self.pivot_data.columns:
            QMessageBox.warning(self, "Warning", f"Element '{selected_element}' not found in pivot data!")
            return

        crm_labels = [label for label in self._inline_crm_rows_display.keys() if 'CRM' in label or 'par' in label]
        if not self.included_crms:
            self.included_crms = {label: QCheckBox(label, checked=True) for label in crm_labels}

        self.current_plot_dialog = PivotPlotDialog(self, selected_element)
        self.current_plot_dialog.exec()