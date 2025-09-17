from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, QLabel, QLineEdit, QPushButton, QMessageBox, QToolTip
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QPen, QColor, QPainter
from PyQt6.QtCharts import QChart, QScatterSeries, QValueAxis, QCategoryAxis,QChartView
import re
import pandas as pd

class PivotPlotDialog(QDialog):
    """Dialog for plotting Verification data with a separate difference plot using PyQt6 QtCharts."""
    def __init__(self, parent, selected_element, annotations):
        super().__init__(parent)
        self.parent = parent
        self.selected_element = selected_element
        self.annotations = annotations
        self.setWindowTitle(f"Verification Plot for {selected_element}")
        self.setGeometry(100, 100, 1400, 900)
        self.setModal(False)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        control_frame = QHBoxLayout()
        self.show_check_crm = QCheckBox("Show Check Verification", checked=True)
        self.show_check_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_check_crm)
        
        self.show_pivot_crm = QCheckBox("Show Pivot Verification", checked=True)
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
        
        correct_btn = QPushButton("Correct Pivot Verification")
        correct_btn.clicked.connect(self.correct_crm_callback)
        control_frame.addWidget(correct_btn)
        
        select_crms_btn = QPushButton("Select Verifications")
        select_crms_btn.clicked.connect(self.open_select_crms_window)
        control_frame.addWidget(select_crms_btn)
        
        report_btn = QPushButton("Report")
        report_btn.clicked.connect(self.show_report)
        control_frame.addWidget(report_btn)
        
        layout.addLayout(control_frame)
        
        self.chart = QChart()
        self.chart_view = QChartView(self.chart)
        self.chart_view.setMouseTracking(True)
        self.chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        layout.addWidget(self.chart_view)
        
        self.diff_chart = QChart()
        self.diff_chart_view = QChartView(self.diff_chart)
        self.diff_chart_view.setMouseTracking(True)
        self.diff_chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        layout.addWidget(self.diff_chart_view)
        
        self.update_plot()

    def show_report(self):
        from .report_dialog import ReportDialog
        dialog = ReportDialog(self, self.annotations)
        dialog.exec()

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
                m = re.search(r'(?i)(?:CRM|par|OREAS)[\s-]*(\d+[a-zA-Z]?)', str(label))
                return m.group(1) if m else None

            crm_groups = {}
            blank_values = {}

            # Step 1: Collect blank rows
            blank_rows = self.parent.pivot_data[
                self.parent.pivot_data['Solution Label'].str.contains(r'CRM\s*BLANK', case=False, na=False, regex=True)
            ]
            print("Blank rows:", blank_rows)

            # Step 2: Get the first blank row's value
            blank_val = 0
            if not blank_rows.empty:
                first_blank_row = blank_rows.iloc[0]
                blank_val = first_blank_row[self.selected_element] if pd.notna(first_blank_row[self.selected_element]) else 0
                blank_val = float(blank_val) if blank_val else 0
            print(f"Selected Blank Value: {blank_val}")

            # Step 3: Assign the same blank value to all crm_ids
            crm_ids = set()
            for sol_label in self.parent._inline_crm_rows_display.keys():
                crm_id = extract_crm_id(sol_label)
                if crm_id:
                    crm_ids.add(crm_id)
                    blank_values[crm_id] = blank_val

            crm_labels = [
                label for label in self.parent._inline_crm_rows_display.keys()
                if ('CRM' in label.upper() or 'par' in label.lower())
                and label not in blank_rows['Solution Label'].values
                and label in self.parent.included_crms and self.parent.included_crms[label].isChecked()
            ]

            y_values = []  # For main chart y-axis range
            diff_values = []  # For difference chart y-axis range

            for sol_label in crm_labels:
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

                                    annotation = f"Verification ID: {crm_id} (Label: {sol_label})"
                                    annotation += f"\n  - Check Verification Value: {crm_val:.4f}"
                                    annotation += f"\n  - Pivot Verification Value: {pivot_val_float:.4f}"
                                    annotation += f"\n  - Original Range: [{lower:.4f} to {upper:.4f}]"
                                    if in_range:
                                        annotation += "\n  - Status: In range (no adjustment needed)."
                                    else:
                                        annotation += "\n  - Status: Out of range without adjustment."

                                    if blank_val != 0:
                                        annotation += f"\n  - Blank Value Subtracted: {blank_val:.4f}"
                                        annotation += f"\n  - Corrected Verification Value: {corrected_crm:.4f}"
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
                                    y_values.extend([crm_val, pivot_val_float, (crm_val + pivot_val_float) / 2, lower, upper])
                                    if crm_val != 0:
                                        diff = ((crm_val - pivot_val_float) / crm_val) * 100
                                        diff_values.append(diff)
                            except ValueError:
                                continue

            if not crm_groups:
                pivot_data = self.parent.pivot_data[
                    self.parent.pivot_data['Solution Label'].str.contains('CRM|par', case=False, na=False)
                ]
                if not pivot_data.empty and self.show_pivot_crm.isChecked():
                    scatter = QScatterSeries()
                    scatter.setName(f"{self.selected_element} (Pivot Verification)")
                    scatter.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                    scatter.setMarkerSize(8)
                    scatter.setPen(QPen(QColor("blue")))
                    sorted_ids = sorted(set(extract_crm_id(row['Solution Label']) for _, row in pivot_data.iterrows() if extract_crm_id(row['Solution Label'])), key=lambda x: (int(re.search(r'\d+', x).group()), x))
                    for i, crm_id in enumerate(sorted_ids):
                        for _, row in pivot_data.iterrows():
                            if extract_crm_id(row['Solution Label']) == crm_id:
                                val = row[self.selected_element]
                                try:
                                    if pd.notna(val):
                                        scatter.append(i + 0.5, float(val))
                                        y_values.append(float(val))
                                except ValueError:
                                    continue
                    if scatter.count() > 0:
                        self.chart.addSeries(scatter)
                        axis_x = QCategoryAxis()
                        for i, crm_id in enumerate(sorted_ids):
                            axis_x.append(f"Verification {crm_id}", i + 0.5)
                        axis_x.setLabelsAngle(45)
                        axis_y = QValueAxis()
                        axis_y.setTitleText(f"{self.selected_element} Value")
                        if y_values:
                            y_min = min(y_values)
                            y_max = max(y_values)
                            margin = (y_max - y_min) * 0.1
                            axis_y.setRange(y_min - margin, y_max + margin)
                        self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
                        self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
                        scatter.attachAxis(axis_x)
                        scatter.attachAxis(axis_y)
                        scatter.hovered.connect(lambda point: QToolTip.showText(
                            self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                            f"Pivot Verification: {point.y():.4f}"
                        ))
                        self.chart.setTitle(f"Pivot Verification Values for {self.selected_element}")
                        self.chart_view.update()
                        self.diff_chart.setTitle("No difference data available")
                        self.diff_chart_view.update()
                        return
                self.chart.setTitle(f"No valid Verification data for {self.selected_element}")
                self.chart_view.update()
                self.diff_chart.setTitle("No difference data available")
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"No valid Verification data for {self.selected_element}")
                return

            sorted_ids = sorted(crm_groups.keys(), key=lambda x: (int(re.search(r'\d+', x).group()), x))
            axis_x = QCategoryAxis()
            for i, crm_id in enumerate(sorted_ids):
                axis_x.append(f"Verification {crm_id}", i + 0.5)
            axis_x.setLabelsAngle(45)
            axis_x.setTitleText("Verification ID")
            self.chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)

            axis_y = QValueAxis()
            axis_y.setTitleText(f"{self.selected_element} Value")
            if y_values:
                y_min = min(y_values)
                y_max = max(y_values)
                margin = (y_max - y_min) * 0.1
                axis_y.setRange(y_min - margin, y_max + margin)
            self.chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)

            axis_x_diff = QCategoryAxis()
            for i, crm_id in enumerate(sorted_ids):
                axis_x_diff.append(f"Verification {crm_id}", i + 0.5)
            axis_x_diff.setLabelsAngle(45)
            axis_x_diff.setTitleText("Verification ID")
            self.diff_chart.addAxis(axis_x_diff, Qt.AlignmentFlag.AlignBottom)

            axis_y_diff = QValueAxis()
            axis_y_diff.setTitleText("Difference (%)")
            axis_y_diff.setLabelFormat("%.1f")
            if diff_values:
                y_min_diff = min(diff_values)
                y_max_diff = max(diff_values)
                margin_diff = (y_max_diff - y_min_diff) * 0.1
                axis_y_diff.setRange(y_min_diff - margin_diff, y_max_diff + margin_diff)
            self.diff_chart.addAxis(axis_y_diff, Qt.AlignmentFlag.AlignLeft)

            plotted = False
            if self.show_check_crm.isChecked():
                scatter_check = QScatterSeries()
                scatter_check.setName(f"{self.selected_element} (Check Verification)")
                scatter_check.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter_check.setMarkerSize(8)
                scatter_check.setPen(QPen(QColor("red")))
                for i, crm_id in enumerate(sorted_ids):
                    for point in crm_groups[crm_id]:
                        if point['crm_val'] != 0:
                            scatter_check.append(i + 0.5, point['crm_val'])
                if scatter_check.count() > 0:
                    self.chart.addSeries(scatter_check)
                    scatter_check.attachAxis(axis_x)
                    scatter_check.attachAxis(axis_y)
                    scatter_check.hovered.connect(lambda point: QToolTip.showText(
                        self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                        f"Check Verification: {point.y():.4f}"
                    ))
                    plotted = True

                if self.show_range.isChecked():
                    scatter_lower = QScatterSeries()
                    scatter_lower.setName("Check Verification Range Lower")
                    scatter_lower.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeTriangle)
                    scatter_lower.setMarkerSize(6)
                    scatter_lower.setPen(QPen(QColor("red")))
                    scatter_upper = QScatterSeries()
                    scatter_upper.setName("Check Verification Range Upper")
                    scatter_upper.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeTriangle)
                    scatter_upper.setMarkerSize(6)
                    scatter_upper.setPen(QPen(QColor("red")))
                    for i, crm_id in enumerate(sorted_ids):
                        for point in crm_groups[crm_id]:
                            range_val = self.parent.calculate_dynamic_range(point['crm_val'])
                            scatter_lower.append(i + 0.5, point['crm_val'] - range_val)
                            scatter_upper.append(i + 0.5, point['crm_val'] + range_val)
                    if scatter_lower.count() > 0:
                        self.chart.addSeries(scatter_lower)
                        self.chart.addSeries(scatter_upper)
                        scatter_lower.attachAxis(axis_x)
                        scatter_lower.attachAxis(axis_y)
                        scatter_upper.attachAxis(axis_x)
                        scatter_upper.attachAxis(axis_y)
                        scatter_lower.hovered.connect(lambda point: QToolTip.showText(
                            self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                            f"Range Lower: {point.y():.4f}"
                        ))
                        scatter_upper.hovered.connect(lambda point: QToolTip.showText(
                            self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                            f"Range Upper: {point.y():.4f}"
                        ))
                        plotted = True

            if self.show_pivot_crm.isChecked():
                scatter_pivot = QScatterSeries()
                scatter_pivot.setName(f"{self.selected_element} (Pivot Verification)")
                scatter_pivot.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter_pivot.setMarkerSize(8)
                for i, crm_id in enumerate(sorted_ids):
                    for point in crm_groups[crm_id]:
                        if point['pivot_val'] != 0:
                            scatter_pivot.append(i + 0.5, point['pivot_val'])
                            scatter_pivot.setPen(QPen(point['color']))
                if scatter_pivot.count() > 0:
                    self.chart.addSeries(scatter_pivot)
                    scatter_pivot.attachAxis(axis_x)
                    scatter_pivot.attachAxis(axis_y)
                    scatter_pivot.hovered.connect(lambda point: QToolTip.showText(
                        self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                        f"Pivot Verification: {point.y():.4f}"
                    ))
                    plotted = True

            if self.show_middle.isChecked():
                scatter_middle = QScatterSeries()
                scatter_middle.setName(f"{self.selected_element} (Middle)")
                scatter_middle.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeRectangle)
                scatter_middle.setMarkerSize(6)
                scatter_middle.setPen(QPen(QColor("green")))
                for i, crm_id in enumerate(sorted_ids):
                    for point in crm_groups[crm_id]:
                        middle_val = (point['crm_val'] + point['pivot_val']) / 2 if point['crm_val'] != 0 and point['pivot_val'] != 0 else point['crm_val'] or point['pivot_val'] or 0
                        if middle_val != 0:
                            scatter_middle.append(i + 0.5, middle_val)
                if scatter_middle.count() > 0:
                    self.chart.addSeries(scatter_middle)
                    scatter_middle.attachAxis(axis_x)
                    scatter_middle.attachAxis(axis_y)
                    scatter_middle.hovered.connect(lambda point: QToolTip.showText(
                        self.chart_view.mapToGlobal(self.chart_view.mapFromScene(point)),
                        f"Middle: {point.y():.4f}"
                    ))
                    plotted = True

            if self.show_diff.isChecked():
                scatter_diff = QScatterSeries()
                scatter_diff.setName("Difference (%)")
                scatter_diff.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
                scatter_diff.setMarkerSize(8)
                scatter_diff.setPen(QPen(QColor("purple")))
                for i, crm_id in enumerate(sorted_ids):
                    for point in crm_groups[crm_id]:
                        if point['crm_val'] != 0:
                            diff = ((point['crm_val'] - point['pivot_val']) / point['crm_val']) * 100
                            scatter_diff.append(i + 0.5, diff)
                if scatter_diff.count() > 0:
                    self.diff_chart.addSeries(scatter_diff)
                    scatter_diff.attachAxis(axis_x_diff)
                    scatter_diff.attachAxis(axis_y_diff)
                    scatter_diff.hovered.connect(lambda point: QToolTip.showText(
                        self.diff_chart_view.mapToGlobal(self.diff_chart_view.mapFromScene(point)),
                        f"Difference: {point.y():.1f}%"
                    ))
                    plotted = True

            if not plotted:
                self.chart.setTitle(f"No data plotted for {self.selected_element}")
                self.diff_chart.setTitle("No difference data plotted")
                self.chart_view.update()
                self.diff_chart_view.update()
                QMessageBox.warning(self, "Warning", f"No data could be plotted for {self.selected_element}")
                return

            self.chart.setTitle(f"Verification Values for {self.selected_element}")
            self.diff_chart.setTitle(f"Difference (%) for {self.selected_element}")
            self.chart_view.update()
            self.diff_chart_view.update()

        except Exception as e:
            self.parent.logger.error(f"Failed to update plot: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to update plot: {str(e)}")

    def open_select_crms_window(self):
        from PyQt6.QtWidgets import QTreeView
        from PyQt6.QtGui import QStandardItemModel, QStandardItem
        w = QDialog(self)
        w.setWindowTitle("Select Verifications to Include")
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