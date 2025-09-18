from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, QLabel, QLineEdit, QPushButton, QMessageBox, QComboBox, QTabWidget
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QColor
import pyqtgraph as pg
import re
import pandas as pd
import numpy as np
import logging

class PivotPlotDialog(QDialog):
    """Dialog for plotting Verification data with PyQtGraph for professional plotting."""
    def __init__(self, parent, selected_element, annotations):
        super().__init__(parent)
        self.parent = parent
        self.selected_element = selected_element
        self.annotations = annotations
        self.setWindowTitle(f"Verification Plot for {selected_element}")
        self.setGeometry(100, 100, 1400, 900)
        self.setModal(False)
        self.range_percent = 5  # Default 5%
        self.logger = logging.getLogger(__name__)
        self.initial_ranges = {}  # To store initial axis ranges for reset
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        control_frame = QHBoxLayout()
        self.show_check_crm = QCheckBox("Show Certificate Value", checked=True)
        self.show_check_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_check_crm)
        
        self.show_pivot_crm = QCheckBox("Show Sample Value", checked=True)
        self.show_pivot_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_pivot_crm)
        
        self.show_middle = QCheckBox("Show Middle", checked=True)
        self.show_middle.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_middle)
        
        self.show_range = QCheckBox("Show Acceptable Range", checked=True)
        self.show_range.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_range)
        
        self.show_diff = QCheckBox("Show Difference", checked=True)
        self.show_diff.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_diff)
        
        # Range percent selection
        control_frame.addWidget(QLabel("Acceptable Range (%):"))
        self.range_combo = QComboBox()
        self.range_combo.addItems(["5%", "10%"])
        self.range_combo.setCurrentText("5%")
        self.range_combo.currentTextChanged.connect(lambda text: setattr(self, 'range_percent', int(text.replace('%', ''))) or self.update_plot())
        control_frame.addWidget(self.range_combo)
        
        control_frame.addWidget(QLabel("Max Correction (%):"))
        self.max_correction_percent = QLineEdit("32")
        control_frame.addWidget(self.max_correction_percent)
        
        correct_btn = QPushButton("Correct Sample Value")
        correct_btn.clicked.connect(self.correct_crm_callback)
        control_frame.addWidget(correct_btn)
        
        select_crms_btn = QPushButton("Select Verifications")
        select_crms_btn.clicked.connect(self.open_select_crms_window)
        control_frame.addWidget(select_crms_btn)
        
        report_btn = QPushButton("Report")
        report_btn.clicked.connect(self.show_report)
        control_frame.addWidget(report_btn)
        
        # Zoom buttons
        zoom_in_btn = QPushButton("Zoom In")
        zoom_in_btn.clicked.connect(self.zoom_in)
        control_frame.addWidget(zoom_in_btn)
        
        zoom_out_btn = QPushButton("Zoom Out")
        zoom_out_btn.clicked.connect(self.zoom_out)
        control_frame.addWidget(zoom_out_btn)
        
        reset_zoom_btn = QPushButton("Reset Zoom")
        reset_zoom_btn.clicked.connect(self.reset_zoom)
        control_frame.addWidget(reset_zoom_btn)
        
        layout.addLayout(control_frame)
        
        # Tab widget for main plot and difference plot
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Main plot tab
        self.main_plot = pg.PlotWidget()
        self.main_plot.setMouseEnabled(x=True, y=True)
        self.main_plot.setMenuEnabled(True)
        self.main_plot.enableAutoRange(x=False, y=False)
        self.main_plot.setBackground('w')
        self.tab_widget.addTab(self.main_plot, "Verification Values")
        
        # Difference plot tab
        self.diff_plot = pg.PlotWidget()
        self.diff_plot.setMouseEnabled(x=True, y=True)
        self.diff_plot.setMenuEnabled(True)
        self.diff_plot.enableAutoRange(x=False, y=False)
        self.diff_plot.setBackground('w')
        self.tab_widget.addTab(self.diff_plot, "Difference (%)")
        
        # Connect mouse move for tooltips
        self.main_plot.scene().sigMouseMoved.connect(self.show_tooltip)
        self.diff_plot.scene().sigMouseMoved.connect(self.show_tooltip)
        
        self.update_plot()

    def zoom_in(self):
        self.main_plot.getViewBox().scaleBy((0.8, 0.8))
        self.diff_plot.getViewBox().scaleBy((0.8, 0.8))

    def zoom_out(self):
        self.main_plot.getViewBox().scaleBy((1.25, 1.25))
        self.diff_plot.getViewBox().scaleBy((1.25, 1.25))

    def reset_zoom(self):
        if self.initial_ranges:
            self.main_plot.setXRange(self.initial_ranges['main_x'][0], self.initial_ranges['main_x'][1])
            self.main_plot.setYRange(self.initial_ranges['main_y'][0], self.initial_ranges['main_y'][1])
            self.diff_plot.setXRange(self.initial_ranges['diff_x'][0], self.initial_ranges['diff_x'][1])
            self.diff_plot.setYRange(self.initial_ranges['diff_y'][0], self.initial_ranges['diff_y'][1])

    def show_tooltip(self, pos):
        plot = self.sender().getPlotItem()
        if not plot:
            return
        vb = plot.getViewBox()
        mouse_point = vb.mapSceneToView(pos)
        for item in plot.listDataItems():
            x, y = item.getData()
            for i in range(len(x)):
                if abs(vb.mapViewToScene(pg.Point(x[i], y[i])).x() - pos.x()) < 10 and abs(vb.mapViewToScene(pg.Point(x[i], y[i])).y() - pos.y()) < 10:
                    plot.clear()
                    self.update_plot()  # Redraw to clear previous tooltip
                    text = pg.TextItem(f"{item.name()}: {self.format_number(y[i])}", anchor=(0, 0))
                    text.setPos(x[i], y[i])
                    plot.addItem(text)
                    return

    def show_report(self):
        from .report_dialog import ReportDialog
        dialog = ReportDialog(self, self.annotations)
        dialog.exec()

    def correct_crm_callback(self):
        self.parent.correct_pivot_crm()
        self.update_plot()

    def is_numeric(self, value):
        """Check if a value can be converted to float."""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False

    def format_number(self, value):
        """Format number to remove trailing zeros."""
        if not self.is_numeric(value):
            return str(value)
        num = float(value)
        if num == 0:
            return "0"
        return f"{num:.4g}"

    def update_plot(self):
        if not self.selected_element or self.selected_element not in self.parent.pivot_data.columns:
            self.logger.warning(f"Element '{self.selected_element}' not found in pivot data!")
            QMessageBox.warning(self, "Warning", f"Element '{self.selected_element}' not found in pivot data!")
            return

        try:
            self.main_plot.clear()
            self.diff_plot.clear()
            self.annotations.clear()

            def extract_crm_id(label):
                m = re.search(r'(?i)(?:CRM|par|OREAS)[\s-]*(\d+[a-zA-Z]?)', str(label))
                return m.group(1) if m else None

            # Extract element name and wavelength
            element_name = self.selected_element.split()[0]
            wavelength = ' '.join(self.selected_element.split()[1:]) if len(self.selected_element.split()) > 1 else ""

            # Get calibration range from Std data
            std_data = self.parent.original_df[
                (self.parent.original_df['Type'] == 'Std') & 
                (self.parent.original_df['Element'] == self.selected_element)
            ]['Soln Conc']
            calibration_min = float(std_data.min()) if not std_data.empty and self.is_numeric(std_data.min()) else 0
            calibration_max = float(std_data.max()) if not std_data.empty and self.is_numeric(std_data.max()) else 0
            calibration_range = f"[{self.format_number(calibration_min)} to {self.format_number(calibration_max)}]" if calibration_min != 0 or calibration_max != 0 else "[0 to 0]"

            # Step 1: Collect blank rows
            blank_rows = self.parent.pivot_data[
                self.parent.pivot_data['Solution Label'].str.contains(r'CRM\s*BLANK', case=False, na=False, regex=True)
            ]
            self.logger.debug(f"Blank rows: {blank_rows}")

            # Step 2: Get the first blank row's value
            blank_val = 0
            if not blank_rows.empty:
                first_blank_row = blank_rows.iloc[0]
                blank_val = first_blank_row[self.selected_element] if pd.notna(first_blank_row[self.selected_element]) else 0
                blank_val = float(blank_val) if self.is_numeric(blank_val) else 0
                if blank_val < 0:
                    self.logger.warning(f"Negative blank value detected: {blank_val}. Setting to 0.")
                    blank_val = 0
            self.logger.debug(f"Selected Blank Value: {blank_val}")

            # Get CRM labels
            crm_labels = [
                label for label in self.parent._inline_crm_rows_display.keys()
                if ('CRM' in label.upper() or 'par' in label.lower())
                and label not in blank_rows['Solution Label'].values
                and label in self.parent.included_crms and self.parent.included_crms[label].isChecked()
            ]

            # Prepare data for plotting
            crm_ids = []
            x_pos = []
            certificate_values = []
            sample_values = []
            corrected_sample_values = []
            lower_bounds = []
            upper_bounds = []
            middle_values = []
            soln_concs = []
            diffs = []

            for sol_label in crm_labels:
                pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                if pivot_row.empty:
                    continue
                pivot_val = pivot_row.iloc[0][self.selected_element]
                crm_id = extract_crm_id(sol_label)
                if not crm_id:
                    continue
                
                # Get Soln Conc for this sample
                sample_rows = self.parent.original_df[
                    (self.parent.original_df['Solution Label'] == sol_label) &
                    (self.parent.original_df['Element'].str.startswith(element_name)) &
                    (self.parent.original_df['Type'].isin(['Samp', 'Sample']))
                ]
                soln_conc = sample_rows['Soln Conc'].iloc[0] if not sample_rows.empty else '---'
                
                for row_data, _ in self.parent._inline_crm_rows_display[sol_label]:
                    if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                        val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                        if not val or not val.strip():
                            continue
                        
                        # Handle non-numeric values
                        if not self.is_numeric(val) or not self.is_numeric(pivot_val):
                            annotation = f"Verification ID: {crm_id} (Label: {sol_label})"
                            annotation += f"\n  - Certificate Value: {val}"
                            annotation += f"\n  - Sample Value: {pivot_val}"
                            annotation += "\n  - Acceptable Range: [N/A]"
                            annotation += "\n  - Status: Out of range (non-numeric data)."
                            annotation += f"\n  - Blank Value Subtracted: {self.format_number(blank_val)}"
                            annotation += f"\n  - Corrected Sample Value: {pivot_val}"
                            annotation += "\n  - Corrected Range: [N/A]"
                            annotation += "\n  - Status after Blank Subtraction: Out of range (non-numeric data)."
                            annotation += f"\n  - Soln Conc: {soln_conc}"
                            annotation += f"\n  - Calibration Range: {calibration_range}"
                            self.annotations.append(annotation)
                            continue

                        try:
                            crm_val = float(val)
                            pivot_val_float = float(pivot_val)
                            
                            # Acceptable range based on selected percent
                            range_val = crm_val * (self.range_percent / 100)
                            lower = crm_val - range_val
                            upper = crm_val + range_val
                            in_range = lower <= pivot_val_float <= upper
                            color = "green" if in_range else "red"

                            # Blank subtraction from sample value
                            corrected_pivot = pivot_val_float - blank_val
                            
                            # Status after blank subtraction
                            corrected_in_range = lower <= corrected_pivot <= upper

                            annotation = f"Verification ID: {crm_id} (Label: {sol_label})"
                            annotation += f"\n  - Certificate Value: {self.format_number(crm_val)}"
                            annotation += f"\n  - Sample Value: {self.format_number(pivot_val_float)}"
                            annotation += f"\n  - Acceptable Range: [{self.format_number(lower)} to {self.format_number(upper)}]"
                            if in_range:
                                annotation += "\n  - Status: In range (no adjustment needed)."
                            else:
                                annotation += "\n  - Status: Out of range without adjustment."

                            if blank_val != 0:
                                annotation += f"\n  - Blank Value Subtracted: {self.format_number(blank_val)}"
                                annotation += f"\n  - Corrected Sample Value: {self.format_number(corrected_pivot)}"
                                annotation += f"\n  - Corrected Range: [{self.format_number(lower)} to {self.format_number(upper)}]"
                                if corrected_in_range:
                                    annotation += "\n  - Status after Blank Subtraction: In range."
                                else:
                                    annotation += "\n  - Status after Blank Subtraction: Out of range."
                                    if corrected_pivot != 0:
                                        if corrected_pivot < lower:
                                            scale_factor = lower / corrected_pivot
                                        elif corrected_pivot > upper:
                                            scale_factor = upper / corrected_pivot
                                        else:
                                            scale_factor = 1.0
                                        scale_percent = abs((scale_factor - 1) * 100)
                                        direction = "increase" if corrected_pivot < lower else "decrease"
                                        annotation += f"\n  - Required Scaling: {scale_percent:.2f}% {direction} to fit within range."
                                        if scale_percent > 200:
                                            color = "darkred"
                                            annotation += f"\n  - Warning: Scaling exceeds 200% ({scale_percent:.2f}%). This point is problematic and may require further investigation."
                                    else:
                                        annotation += "\n  - Scaling not applicable (corrected sample value is zero)."
                            else:
                                annotation += "\n  - No blank subtraction applied (Blank Value: 0)."
                            
                            # Add Soln Conc and Calibration Range
                            annotation += f"\n  - Soln Conc: {soln_conc if isinstance(soln_conc, str) else self.format_number(soln_conc)}"
                            annotation += f"\n  - Calibration Range: {calibration_range}"

                            self.annotations.append(annotation)
                            
                            # Add to plotting data
                            crm_ids.append(crm_id)
                            x_pos.append(len(crm_ids) - 0.5)
                            certificate_values.append(crm_val)
                            sample_values.append(pivot_val_float)
                            corrected_sample_values.append(corrected_pivot)
                            lower_bounds.append(lower)
                            upper_bounds.append(upper)
                            middle_values.append((crm_val + pivot_val_float) / 2)
                            soln_concs.append(soln_conc)
                            if crm_val != 0:
                                diffs.append(((crm_val - corrected_pivot) / crm_val) * 100)
                            else:
                                diffs.append(0)
                        except ValueError as e:
                            self.logger.error(f"ValueError in processing data for {sol_label}: {str(e)}")
                            continue

            if not crm_ids:
                self.main_plot.clear()
                self.diff_plot.clear()
                QMessageBox.warning(self, "Warning", f"No valid Verification data for {self.selected_element}")
                return

            # Set up main plot
            self.main_plot.setLabel('bottom', 'Verification ID')
            self.main_plot.setLabel('left', f'{self.selected_element} Value')
            self.main_plot.setTitle(f'Verification Values for {self.selected_element}')
            self.main_plot.getAxis('bottom').setTicks([[(i - 0.5, f'V {id}') for i, id in enumerate(crm_ids)]])
            y_values = certificate_values + sample_values + corrected_sample_values + lower_bounds + upper_bounds + middle_values
            if y_values:
                y_min, y_max = min(y_values), max(y_values)
                margin = (y_max - y_min) * 0.1
                self.main_plot.setXRange(-0.5, len(crm_ids) - 0.5)
                self.main_plot.setYRange(y_min - margin, y_max + margin)
                self.initial_ranges['main_x'] = (-0.5, len(crm_ids) - 0.5)
                self.initial_ranges['main_y'] = (y_min - margin, y_max + margin)

            # Plot main chart
            if self.show_check_crm.isChecked():
                scatter = pg.PlotDataItem(x=x_pos, y=certificate_values, pen=None, symbol='o', symbolSize=8, symbolPen='r', name='Certificate Value')
                self.main_plot.addItem(scatter)

            if self.show_pivot_crm.isChecked():
                for i in range(len(crm_ids)):
                    color = 'g' if lower_bounds[i] <= sample_values[i] <= upper_bounds[i] else 'r'
                    scatter = pg.PlotDataItem(x=[x_pos[i]], y=[sample_values[i]], pen=None, symbol='o', symbolSize=8, symbolPen=color, name='Sample Value')
                    self.main_plot.addItem(scatter)

            if self.show_middle.isChecked():
                scatter = pg.PlotDataItem(x=x_pos, y=middle_values, pen=None, symbol='s', symbolSize=6, symbolPen='g', name='Middle')
                self.main_plot.addItem(scatter)

            if self.show_range.isChecked():
                for i in range(len(crm_ids)):
                    line_lower = pg.PlotDataItem(x=[x_pos[i] - 0.2, x_pos[i] + 0.2], y=[lower_bounds[i], lower_bounds[i]], pen=pg.mkPen('c', width=2))
                    line_upper = pg.PlotDataItem(x=[x_pos[i] - 0.2, x_pos[i] + 0.2], y=[upper_bounds[i], upper_bounds[i]], pen=pg.mkPen('c', width=2))
                    self.main_plot.addItem(line_lower)
                    self.main_plot.addItem(line_upper)

            self.main_plot.showGrid(x=True, y=True, alpha=0.3)

            # Set up difference plot
            self.diff_plot.setLabel('bottom', 'Verification ID')
            self.diff_plot.setLabel('left', 'Difference (%)')
            self.diff_plot.setTitle(f'Difference (%) for {self.selected_element}')
            self.diff_plot.getAxis('bottom').setTicks([[(i - 0.5, f'V {id}') for i, id in enumerate(crm_ids)]])
            if diffs:
                y_min, y_max = min(diffs), max(diffs)
                margin = (y_max - y_min) * 0.1
                self.diff_plot.setXRange(-0.5, len(crm_ids) - 0.5)
                self.diff_plot.setYRange(y_min - margin, y_max + margin)
                self.initial_ranges['diff_x'] = (-0.5, len(crm_ids) - 0.5)
                self.initial_ranges['diff_y'] = (y_min - margin, y_max + margin)

            if self.show_diff.isChecked():
                scatter = pg.PlotDataItem(x=x_pos, y=diffs, pen=None, symbol='o', symbolSize=8, symbolPen='m', name='Difference (%)')
                self.diff_plot.addItem(scatter)
                self.diff_plot.addLine(y=0, pen=pg.mkPen('k', style=Qt.PenStyle.DashLine))

            self.diff_plot.showGrid(x=True, y=True, alpha=0.3)

        except Exception as e:
            self.logger.error(f"Failed to update plot: {str(e)}", exc_info=True)
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