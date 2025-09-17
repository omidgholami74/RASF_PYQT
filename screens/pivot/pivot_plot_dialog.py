from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, QLabel, QLineEdit, QPushButton, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPen, QColor,QPainter
from PyQt6.QtCharts import QChart, QLineSeries, QScatterSeries, QValueAxis, QCategoryAxis,QChartView
import re
import numpy as np
import pandas as pd

class ZoomableChartView(QChartView):
    """Custom QChartView with mouse wheel zoom."""
    def __init__(self, chart, parent=None):
        super().__init__(chart, parent)
        self.setMouseTracking(True)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)

    def wheelEvent(self, event):
        factor = 1.1 if event.angleDelta().y() > 0 else 1 / 1.1
        self.chart().zoom(factor)

class PivotPlotDialog(QDialog):
    """Dialog for plotting CRM data with a separate difference plot using PyQt6 QtCharts."""
    def __init__(self, parent, selected_element, annotations):
        super().__init__(parent)
        self.parent = parent
        self.selected_element = selected_element
        self.annotations = annotations
        self.setWindowTitle(f"CRM Plot for {selected_element}")
        self.setGeometry(100, 100, 1400, 900)
        self.setWindowState(Qt.WindowState.WindowFullScreen)
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
        
        report_btn = QPushButton("Report")
        report_btn.clicked.connect(self.show_report)
        control_frame.addWidget(report_btn)
        
        zoom_out_btn = QPushButton("Zoom Out")
        zoom_out_btn.clicked.connect(self.zoom_out)
        control_frame.addWidget(zoom_out_btn)
        
        layout.addLayout(control_frame)
        
        self.chart = QChart()
        self.chart_view = ZoomableChartView(self.chart)
        layout.addWidget(self.chart_view)
        
        self.diff_chart = QChart()
        self.diff_chart_view = ZoomableChartView(self.diff_chart)
        layout.addWidget(self.diff_chart_view)
        
        self.update_plot()

    def show_report(self):
        from report_dialog import ReportDialog
        dialog = ReportDialog(self, self.annotations)
        dialog.exec()

    def zoom_out(self):
        self.chart.zoomReset()
        self.diff_chart.zoomReset()

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

            self.annotations.clear()

            def extract_crm_id(label):
                m = re.search(r'(\d+)', str(label).replace(' ', ''), re.IGNORECASE)
                return m.group(1) if m else None

            crm_groups = {}
            blank_values = {}

            blank_rows = self.parent.pivot_data[self.parent.pivot_data['Solution Label'].str.contains(r'CRM\s*BLANK', case=False, na=False, regex=True)]
            for _, row in blank_rows.iterrows():
                bid = extract_crm_id(row['Solution Label'])
                if bid:
                    blank_val = row[self.selected_element] if pd.notna(row[self.selected_element]) else 0
                    blank_values[bid] = blank_val

            crm_labels = [label for label in self.parent._inline_crm_rows_display.keys() 
                          if ('CRM' in label.upper() or 'par' in label.lower()) 
                          and label not in blank_rows['Solution Label'].values]

            for sol_label in crm_labels:
                if sol_label not in self.parent.included_crms or not self.parent.included_crms[sol_label].isChecked():
                    continue
                pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                if pivot_row.empty:
                    continue
                pivot_val = pivot_row.iloc[0][self.selected_element]
                crm_id = extract_crm_id(sol_label)
                if not crm_id:
                    continue
                if crm_id not in crm_groups:
                    crm_groups[crm_id] = []
                for row_data, _ in self.parent._inline_crm_rows_display[sol_label]:
                    if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                        val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                        if val and val.strip():
                            try:
                                crm_val = float(val)
                                if pd.notna(pivot_val):
                                    pivot_val_float = float(pivot_val)
                                    range_val = self.parent.calculate_dynamic_range(crm_val)
                                    lower = crm_val - range_val
                                    upper = crm_val + range_val
                                    in_range = lower <= pivot_val_float <= upper
                                    color = QColor("green") if in_range else QColor("red")
                                    
                                    blank_val = blank_values.get(crm_id, 0)
                                    corrected_crm = crm_val - blank_val
                                    corrected_range_val = self.parent.calculate_dynamic_range(corrected_crm)
                                    corrected_lower = corrected_crm - corrected_range_val
                                    corrected_upper = corrected_crm + corrected_range_val
                                    corrected_in_range = corrected_lower <= pivot_val_float <= corrected_upper
                                    
                                    annotation = f"CRM ID: {crm_id} (Label: {sol_label})"
                                    annotation += f"\n  - Check CRM Value: {crm_val:.4f}"
                                    annotation += f"\n  - Pivot CRM Value: {pivot_val_float:.4f}"
                                    annotation += f"\n  - Original Range: [{lower:.4f} to {upper:.4f}]"
                                    if in_range:
                                        annotation += "\n  - Status: In range (no adjustment needed)."
                                    else:
                                        annotation += "\n  - Status: Out of range without adjustment."
                                    
                                    if blank_val != 0:
                                        annotation += f"\n  - Blank Value Subtracted: {blank_val:.4f}"
                                        annotation += f"\n  - Corrected CRM Value: {corrected_crm:.4f}"
                                        annotation += f"\n  - Corrected Range: [{corrected_lower:.4f} to {corrected_upper:.4f}]"
                                        if corrected_in_range:
                                            annotation += "\n  - Status after Blank Subtraction: In range."
                                        else:
                                            annotation += "\n  - Status after Blank Subtraction: Out of range."
                                            if pivot_val_float != 0:
                                                if pivot_val_float < corrected_lower:
                                                    scale_factor = corrected_lower / pivot_val_float
                                                elif pivot_val_float > corrected_upper:
                                                    scale_factor = corrected_upper / pivot_val_float
                                                else:
                                                    scale_factor = 1.0
                                                scale_percent = (scale_factor - 1) * 100 if pivot_val_float < corrected_lower else (1 - scale_factor) * 100
                                                direction = "increase" if pivot_val_float < corrected_lower else "decrease"
                                                annotation += f"\n  - Required Scaling: {abs(scale_percent):.2f}% {direction} to fit within range."
                                                annotation += f"\n  - Applying this scaling factor would bring the value into the acceptable range."
                                                if abs(scale_percent) > 200:
                                                    color = QColor("darkred")
                                                    annotation += f"\n  - Warning: Scaling exceeds 200% ({abs(scale_percent):.2f}%). This point is problematic and may require further investigation."
                                            else:
                                                annotation += "\n  - Scaling not applicable (pivot value is zero)."
                                    else:
                                        annotation += "\n  - No blank subtraction applied (Blank Value: 0)."
                                    
                                    self.annotations.append(annotation)
                                    
                                    crm_groups[crm_id].append({
                                        'label': sol_label,
                                        'crm_val': crm_val,
                                        'pivot_val': pivot_val_float,
                                        'color': color,
                                        'annotation': annotation,
                                        'blank_sub': blank_val != 0
                                    })
                            except ValueError:
                                continue

            if not crm_groups:
                pivot_data = self.parent.pivot_data[self.parent.pivot_data['Solution Label'].str.contains('CRM|par', case=False, na=False)]
                if not pivot_data.empty and self.show_pivot_crm.isChecked():
                    series = QLineSeries()
                    series.setName(f"{self.selected_element} (Pivot CRM)")
                    series.setPen(QPen(QColor("blue"), 2))
                    for i, (_, row) in enumerate(pivot_data.iterrows()):
                        val = row[self.selected_element]
                        try:
                            if pd.notna(val):
                                series.append(i + 0.5, float(val))
                        except ValueError:
                            continue
                    if series.count() > 0:
                        self.chart.addSeries(series)
                        axis_x = QCategoryAxis()
                        for i, (_, row) in enumerate(pivot_data.iterrows()):
                            axis_x.append(extract_crm_id(row['Solution Label']) or row['Solution Label'], i + 0.5)
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

            sorted_ids = sorted(crm_groups.keys(), key=lambda x: int(x))
            labels = [f"CRM {id}" for id in sorted_ids]
            positions = list(range(len(sorted_ids)))

            agg_crm_vals = []
            agg_pivot_vals = []
            agg_colors = []
            agg_annotations = []

            for pos, cid in enumerate(sorted_ids):
                group = crm_groups[cid]
                if not group:
                    continue
                avg_crm = np.mean([p['crm_val'] for p in group])
                avg_pivot = np.mean([p['pivot_val'] for p in group])
                color = group[0]['color']
                annotation = group[0]['annotation']
                agg_crm_vals.append(avg_crm)
                agg_pivot_vals.append(avg_pivot)
                agg_colors.append(color)
                agg_annotations.append(annotation)

            if not agg_crm_vals:
                self.chart.setTitle(f"No valid data for {self.selected_element}")
                self.chart_view.update()
                self.diff_chart.setTitle("No difference data available")
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"No valid data for {self.selected_element}")
                return

            middle_values = [(disp_val + piv_val) / 2 if disp_val != 0 and piv_val != 0 else disp_val or piv_val or 0
                             for disp_val, piv_val in zip(agg_crm_vals, agg_pivot_vals)]
            range_lower = [val - self.parent.calculate_dynamic_range(val) for val in agg_crm_vals]
            range_upper = [val + self.parent.calculate_dynamic_range(val) for val in agg_crm_vals]
            diff_values = [((crm_val - piv_val) / crm_val) * 100 if crm_val != 0 else 0
                           for crm_val, piv_val in zip(agg_crm_vals, agg_pivot_vals)]

            axis_x = QCategoryAxis()
            for i, label in enumerate(labels):
                axis_x.append(label, i + 0.5)
            axis_x.setLabelsAngle(45)
            axis_x.setTitleText("CRM ID")
            self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)

            axis_y = QValueAxis()
            axis_y.setTitleText(f"{self.selected_element} Value")
            self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)

            axis_x_diff = QCategoryAxis()
            for i, label in enumerate(labels):
                axis_x_diff.append(label, i + 0.5)
            axis_x_diff.setLabelsAngle(45)
            axis_x_diff.setTitleText("CRM ID")
            self.diff_chart.addAxis(axis_x_diff, Qt.AlignmentFlag.AlignBottom)

            axis_y_diff = QValueAxis()
            axis_y_diff.setTitleText("Difference (%)")
            axis_y_diff.setLabelFormat("%.1f")
            self.diff_chart.addAxis(axis_y_diff, Qt.AlignmentFlag.AlignLeft)

            plotted = False
            if self.show_check_crm.isChecked() and any(v != 0 for v in agg_crm_vals):
                series = QLineSeries()
                series.setName(f"{self.selected_element} (Check CRM)")
                series.setPen(QPen(QColor("red"), 2, Qt.PenStyle.DashLine))
                for i, val in enumerate(agg_crm_vals):
                    series.append(i + 0.5, val)
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
                        range_series_lower.append(i + 0.5, lower)
                    for i, upper in enumerate(range_upper):
                        range_series_upper.append(i + 0.5, upper)
                    self.chart.addSeries(range_series_lower)
                    self.chart.addSeries(range_series_upper)
                    range_series_lower.attachAxis(axis_x)
                    range_series_lower.attachAxis(axis_y)
                    range_series_upper.attachAxis(axis_x)
                    range_series_upper.attachAxis(axis_y)
                plotted = True

            if self.show_pivot_crm.isChecked() and any(v != 0 for v in agg_pivot_vals):
                series = QLineSeries()
                series.setName(f"{self.selected_element} (Pivot CRM)")
                series.setPen(QPen(QColor("blue"), 2))
                scatter = QScatterSeries()
                scatter.setName("")
                scatter.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter.setMarkerSize(8)
                for i, (val, color) in enumerate(zip(agg_pivot_vals, agg_colors)):
                    series.append(i + 0.5, val)
                    scatter.append(i + 0.5, val)
                    scatter.setPen(QPen(color))
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
                    series.append(i + 0.5, val)
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
                    series.append(i + 0.5, val)
                    scatter.append(i + 0.5, val)
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

            self.chart.setTitle(f"CRM Values for {self.selected_element} (Grouped by CRM ID)")
            self.diff_chart.setTitle(f"Difference (%) for {self.selected_element}")
            self.chart_view.update()
            self.diff_chart_view.update()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to update plot: {str(e)}")

    def open_select_crms_window(self):
        from PyQt6.QtWidgets import QTreeView
        from PyQt6.QtGui import QStandardItemModel, QStandardItem
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