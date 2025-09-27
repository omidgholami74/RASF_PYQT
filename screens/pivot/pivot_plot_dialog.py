from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, QLabel, QLineEdit, QPushButton, QMessageBox, QComboBox, QTreeView
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import pyqtgraph as pg
import re
import pandas as pd
import numpy as np
import logging
from datetime import datetime

class PivotPlotDialog(QDialog):
    """Dialog for plotting Verification data with PyQtGraph."""
    def __init__(self, parent, annotations):
        super().__init__(parent)
        self.parent = parent
        self.selected_element = ""
        self.annotations = annotations
        self.setWindowTitle("Verification Plot")
        self.setGeometry(100, 100, 1400, 900)
        self.setModal(False)
        self.range_percent = 5
        self.logger = logging.getLogger(__name__)
        self.initial_ranges = {}
        self.element_selector = QComboBox()
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        control_frame = QHBoxLayout()
        control_frame.addWidget(QLabel("Select Element:"))
        if self.parent.pivot_data is not None:
            columns = [col for col in self.parent.pivot_data.columns if col != 'Solution Label']
            self.element_selector.addItems(columns)
        self.element_selector.currentTextChanged.connect(self.update_plot)
        control_frame.addWidget(self.element_selector)
        
        self.show_check_crm = QCheckBox("Show Certificate Value", checked=True)
        self.show_check_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_check_crm)
        
        self.show_pivot_crm = QCheckBox("Show Sample Value", checked=True)
        self.show_pivot_crm.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_pivot_crm)
        
        self.show_range = QCheckBox("Show Acceptable Range", checked=True)
        self.show_range.toggled.connect(self.update_plot)
        control_frame.addWidget(self.show_range)
        
        control_frame.addWidget(QLabel("Acceptable Range (% for >100):"))
        self.range_combo = QComboBox()
        self.range_combo.addItems(["5%", "10%", "15%"])
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
        
        self.main_plot = pg.PlotWidget()
        self.main_plot.setMouseEnabled(x=True, y=True)
        self.main_plot.setMenuEnabled(True)
        self.main_plot.enableAutoRange(x=False, y=False)
        self.main_plot.setBackground('w')
        self.legend = self.main_plot.addLegend(offset=(10, 10))
        layout.addWidget(self.main_plot)
        
        self.main_plot.scene().sigMouseMoved.connect(self.show_tooltip)
        
        self.update_plot()

    def zoom_in(self):
        self.main_plot.getViewBox().scaleBy((0.8, 0.8))

    def zoom_out(self):
        self.main_plot.getViewBox().scaleBy((1.25, 1.25))

    def reset_zoom(self):
        if self.initial_ranges:
            self.main_plot.setXRange(self.initial_ranges['main_x'][0], self.initial_ranges['main_x'][1])
            self.main_plot.setYRange(self.initial_ranges['main_y'][0], self.initial_ranges['main_y'][1])

    def show_tooltip(self, pos):
        plot = self.main_plot.getPlotItem()
        if not plot:
            return
        vb = plot.getViewBox()
        mouse_point = vb.mapSceneToView(pos)
        for item in plot.items[:]:
            if isinstance(item, pg.TextItem):
                plot.removeItem(item)
        for item in plot.listDataItems():
            x, y = item.getData()
            for i in range(len(x)):
                if abs(vb.mapViewToScene(pg.Point(x[i], y[i])).x() - pos.x()) < 20 and abs(vb.mapViewToScene(pg.Point(x[i], y[i])).y() - pos.y()) < 20:
                    text = pg.TextItem(f"{item.name()}: {self.format_number(y[i])}", anchor=(0, 0))
                    text.setPos(x[i], y[i])
                    plot.addItem(text)
                    return

    def show_report(self):
        from .report_dialog import ReportDialog
        self.logger.debug(f"Opening report with annotations: {self.annotations}")
        dialog = ReportDialog(self, self.annotations)
        dialog.exec()

    def correct_crm_callback(self):
        self.parent.correct_pivot_crm()
        self.update_plot()

    def is_numeric(self, value):
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False

    def format_number(self, value):
        if not self.is_numeric(value):
            return str(value)
        num = float(value)
        if num == 0:
            return "0"
        return f"{num:.4f}".rstrip('0').rstrip('.')

    def update_plot(self):
        self.selected_element = self.element_selector.currentText()
        if not self.selected_element or self.selected_element not in self.parent.pivot_data.columns:
            self.logger.warning(f"Element '{self.selected_element}' not found in pivot data!")
            QMessageBox.warning(self, "Warning", f"Element '{self.selected_element}' not found!")
            return

        try:
            self.main_plot.clear()
            self.legend.clear()
            self.annotations.clear()
            added_legend_names = set()

            def extract_crm_id(label):
                m = re.search(r'(?i)(?:\bCRM\b|\bOREAS\b)?[\s-]*(\d+[a-zA-Z]?)[\s-]*(?:\bpar\b)?', str(label))
                return m.group(1) if m else str(label)

            element_name = self.selected_element.split()[0]
            wavelength = ' '.join(self.selected_element.split()[1:]) if len(self.selected_element.split()) > 1 else ""
            analysis_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            std_data = self.parent.original_df[
                (self.parent.original_df['Type'] == 'Std') & 
                (self.parent.original_df['Element'] == self.selected_element)
            ]['Soln Conc']
            std_data_numeric = [float(x) for x in std_data if self.is_numeric(x)]
            if not std_data_numeric:
                self.logger.warning(f"No valid Std data for {self.selected_element}")
                calibration_min = 0
                calibration_max = 0
                calibration_range = "[0 to 0]"
            else:
                calibration_min = min(std_data_numeric)
                calibration_max = max(std_data_numeric)
                calibration_range = f"[{self.format_number(calibration_min)} to {self.format_number(calibration_max)}]"

            sample_data = self.parent.original_df[
                (self.parent.original_df['Type'].isin(['Samp', 'Sample'])) &
                (self.parent.original_df['Element'] == self.selected_element)
            ]['Soln Conc']
            sample_data_numeric = [float(x) for x in sample_data if self.is_numeric(x)]
            if not sample_data_numeric:
                self.logger.warning(f"No valid Sample data for {self.selected_element}")
                soln_conc_min = '---'
                soln_conc_max = '---'
                soln_conc_range = '---'
                in_calibration_range_soln = False
            else:
                soln_conc_min = min(sample_data_numeric)
                soln_conc_max = max(sample_data_numeric)
                soln_conc_range = f"[{self.format_number(soln_conc_min)} to {self.format_number(soln_conc_max)}]"
                in_calibration_range_soln = (
                    calibration_min <= soln_conc_min <= calibration_max and
                    calibration_min <= soln_conc_max <= calibration_max
                ) if calibration_min != 0 or calibration_max != 0 else False

            blank_rows = self.parent.pivot_data[
                self.parent.pivot_data['Solution Label'].str.contains(r'CRM\s*BLANK', case=False, na=False, regex=True)
            ]
            blank_val = 0
            blank_correction_status = "Not Applied"
            selected_blank_label = "None"
            if not blank_rows.empty:
                best_blank_val = 0
                best_blank_label = "None"
                min_distance = float('inf')
                in_range_found = False
                for _, row in blank_rows.iterrows():
                    candidate_blank = row[self.selected_element] if pd.notna(row[self.selected_element]) else 0
                    candidate_label = row['Solution Label']
                    if not self.is_numeric(candidate_blank):
                        self.logger.debug(f"Skipping non-numeric blank value '{candidate_blank}' for label '{candidate_label}'")
                        continue
                    candidate_blank = float(candidate_blank)
                    in_range = False
                    for sol_label in self.parent._inline_crm_rows_display.keys():
                        if sol_label in blank_rows['Solution Label'].values:
                            continue
                        pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                        if pivot_row.empty:
                            continue
                        pivot_val = pivot_row.iloc[0][self.selected_element]
                        if not self.is_numeric(pivot_val):
                            continue
                        pivot_val_float = float(pivot_val)
                        for row_data, _ in self.parent._inline_crm_rows_display[sol_label]:
                            if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                                val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                                if not self.is_numeric(val):
                                    continue
                                crm_val = float(val)
                                range_val = self.parent.calculate_dynamic_range(crm_val)
                                lower, upper = crm_val - range_val, crm_val + range_val
                                corrected_pivot = pivot_val_float - candidate_blank
                                if lower <= corrected_pivot <= upper:
                                    in_range = True
                                    break
                        if in_range:
                            break
                    if in_range:
                        best_blank_val = candidate_blank
                        best_blank_label = candidate_label
                        in_range_found = True
                        break
                    if not in_range_found:
                        for sol_label in self.parent._inline_crm_rows_display.keys():
                            if sol_label in blank_rows['Solution Label'].values:
                                continue
                            pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                            if pivot_row.empty:
                                continue
                            pivot_val = pivot_row.iloc[0][self.selected_element]
                            if not self.is_numeric(pivot_val):
                                continue
                            pivot_val_float = float(pivot_val)
                            for row_data, _ in self.parent._inline_crm_rows_display[sol_label]:
                                if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                                    val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                                    if not self.is_numeric(val):
                                        continue
                                    crm_val = float(val)
                                    corrected_pivot = pivot_val_float - candidate_blank
                                    distance = abs(corrected_pivot - crm_val)
                                    if distance < min_distance:
                                        min_distance = distance
                                        best_blank_val = candidate_blank
                                        best_blank_label = candidate_label
                blank_val = best_blank_val
                selected_blank_label = best_blank_label
                blank_correction_status = "Applied" if blank_val != 0 else "Not Applied"

            crm_labels = [
                label for label in self.parent._inline_crm_rows_display.keys()
                if label not in blank_rows['Solution Label'].values
                and label in self.parent.included_crms and self.parent.included_crms[label].isChecked()
            ]

            crm_id_to_labels = {}
            for sol_label in crm_labels:
                crm_id = extract_crm_id(sol_label)
                if crm_id not in crm_id_to_labels:
                    crm_id_to_labels[crm_id] = []
                crm_id_to_labels[crm_id].append(sol_label)

            unique_crm_ids = sorted(crm_id_to_labels.keys())
            x_pos_map = {crm_id: i for i, crm_id in enumerate(unique_crm_ids)}
            certificate_values = {}
            sample_values = {}
            corrected_sample_values = {}
            lower_bounds = {}
            upper_bounds = {}
            soln_concs = {}
            int_values = {}

            for crm_id in unique_crm_ids:
                certificate_values[crm_id] = []
                sample_values[crm_id] = []
                corrected_sample_values[crm_id] = []
                lower_bounds[crm_id] = []
                upper_bounds[crm_id] = []
                soln_concs[crm_id] = []
                int_values[crm_id] = []

                for sol_label in crm_id_to_labels[crm_id]:
                    pivot_row = self.parent.pivot_data[self.parent.pivot_data['Solution Label'] == sol_label]
                    if pivot_row.empty:
                        self.logger.warning(f"No pivot data for {sol_label}")
                        continue
                    pivot_val = pivot_row.iloc[0][self.selected_element]
                    if pd.isna(pivot_val) or not self.is_numeric(pivot_val):
                        self.logger.warning(f"Invalid pivot value for {sol_label}: {pivot_val}")
                        pivot_val = 0
                    else:
                        pivot_val = float(pivot_val)

                    sample_rows = self.parent.original_df[
                        (self.parent.original_df['Solution Label'] == sol_label) &
                        (self.parent.original_df['Element'].str.startswith(element_name)) &
                        (self.parent.original_df['Type'].isin(['Samp', 'Sample']))
                    ]
                    soln_conc = sample_rows['Soln Conc'].iloc[0] if not sample_rows.empty else '---'
                    int_val = sample_rows['Int'].iloc[0] if not sample_rows.empty else '---'
                    
                    int_values_list = sample_rows['Int'].dropna().astype(float).tolist()
                    rsd_percent = (np.std(int_values_list) / np.mean(int_values_list) * 100) if int_values_list and np.mean(int_values_list) != 0 else 0.0
                    
                    detection_limit = 0.2
                    crm_source = "NIST"
                    sample_matrix = "Soil"
                    
                    for row_data, _ in self.parent._inline_crm_rows_display[sol_label]:
                        if isinstance(row_data, list) and row_data and row_data[0].endswith("CRM"):
                            val = row_data[self.parent.pivot_data.columns.get_loc(self.selected_element)] if self.selected_element in self.parent.pivot_data.columns else ""
                            if not val or not self.is_numeric(val):
                                self.logger.warning(f"Invalid CRM value for {sol_label}: {val}")
                                annotation = f"Verification ID: {crm_id} (Label: {sol_label})"
                                annotation += f"\n  - Certificate Value: {val or 'N/A'}"
                                annotation += f"\n  - Sample Value: {self.format_number(pivot_val)}"
                                annotation += f"\n  - Acceptable Range: [N/A]"
                                annotation += f"\n  - Status: Out of range (non-numeric data)."
                                annotation += f"\n  - Blank Value: {self.format_number(blank_val)}"
                                annotation += f"\n  - Blank Label: {selected_blank_label}"
                                annotation += f"\n  - Blank Correction Status: {blank_correction_status}"
                                annotation += f"\n  - Sample Value - Blank: {self.format_number(pivot_val)}"
                                annotation += f"\n  - Corrected Range: [N/A]"
                                annotation += f"\n  - Status after Blank Subtraction: Out of range (non-numeric data)."
                                annotation += f"\n  - Soln Conc: {soln_conc if isinstance(soln_conc, str) else self.format_number(soln_conc)} {'in_range' if in_calibration_range_soln else 'out_range'}"
                                annotation += f"\n  - Int: {int_val if isinstance(int_val, str) else self.format_number(int_val)}"
                                annotation += f"\n  - Calibration Range: {calibration_range} {'in_range' if in_calibration_range_soln else 'out_range'}"
                                annotation += f"\n  - CRM Source: {crm_source}"
                                annotation += f"\n  - Sample Matrix: {sample_matrix}"
                                annotation += f"\n  - Element Wavelength: {wavelength}"
                                annotation += f"\n  - Analysis Date: {analysis_date}"
                                self.annotations.append(annotation)
                                continue

                            crm_val = float(val)
                            pivot_val_float = float(pivot_val)
                            range_val = self.parent.calculate_dynamic_range(crm_val)
                            lower = crm_val - range_val
                            upper = crm_val + range_val
                            in_range = lower <= pivot_val_float <= upper

                            annotation = f"Verification ID: {crm_id} (Label: {sol_label})"
                            annotation += f"\n  - Certificate Value: {self.format_number(crm_val)}"
                            annotation += f"\n  - Sample Value: {self.format_number(pivot_val_float)}"
                            annotation += f"\n  - Acceptable Range: [{self.format_number(lower)} to {self.format_number(upper)}]"
                            corrected_in_range = False
                            if in_range:
                                annotation += f"\n  - Status: In range (no adjustment needed)."
                                corrected_pivot = pivot_val_float
                                corrected_in_range = True
                                annotation += f"\n  - Blank Value: {self.format_number(blank_val)}"
                                annotation += f"\n  - Blank Label: {selected_blank_label}"
                                annotation += f"\n  - Blank Correction Status: Not Applied (in range)"
                                annotation += f"\n  - Sample Value - Blank: {self.format_number(corrected_pivot)}"
                                annotation += f"\n  - Status after Blank Subtraction: In range."
                            else:
                                annotation += f"\n  - Status: Out of range without adjustment."
                                corrected_pivot = pivot_val_float - blank_val
                                annotation += f"\n  - Blank Value: {self.format_number(blank_val)}"
                                annotation += f"\n  - Blank Label: {selected_blank_label}"
                                annotation += f"\n  - Blank Correction Status: {blank_correction_status}"
                                annotation += f"\n  - Sample Value - Blank: {self.format_number(corrected_pivot)}"
                                annotation += f"\n  - Corrected Range: [{self.format_number(lower)} to {self.format_number(upper)}]"
                                corrected_in_range = lower <= corrected_pivot <= upper
                                if corrected_in_range:
                                    annotation += f"\n  - Status after Blank Subtraction: In range."
                                else:
                                    annotation += f"\n  - Status after Blank Subtraction: Out of range."
                                    if corrected_pivot != 0:
                                        if corrected_pivot < lower:
                                            scale_factor = lower / corrected_pivot
                                            direction = "increase"
                                        elif corrected_pivot > upper:
                                            scale_factor = upper / corrected_pivot
                                            direction = "decrease"
                                        else:
                                            scale_factor = 1.0
                                            direction = ""
                                        scale_percent = abs((scale_factor - 1) * 100)
                                        annotation += f"\n  - Required Scaling: {scale_percent:.2f}% {direction} to fit within range."
                                        if scale_percent > float(self.max_correction_percent.text()):
                                            annotation += f"\n  - Warning: Scaling exceeds {self.max_correction_percent.text()}% ({scale_percent:.2f}%)."
                                    else:
                                        annotation += f"\n  - Scaling not applicable (corrected sample value is zero)."
                            
                            annotation += f"\n  - Soln Conc: {soln_conc if isinstance(soln_conc, str) else self.format_number(soln_conc)} {'in_range' if in_calibration_range_soln else 'out_range'}"
                            annotation += f"\n  - Int: {int_val if isinstance(int_val, str) else self.format_number(int_val)}"
                            annotation += f"\n  - Calibration Range: {calibration_range} {'in_range' if in_calibration_range_soln else 'out_range'}"
                            annotation += f"\n  - CRM Source: {crm_source}"
                            annotation += f"\n  - Sample Matrix: {sample_matrix}"
                            annotation += f"\n  - Element Wavelength: {wavelength}"
                            annotation += f"\n  - Analysis Date: {analysis_date}"

                            self.annotations.append(annotation)
                            
                            certificate_values[crm_id].append(crm_val)
                            sample_values[crm_id].append(pivot_val_float)
                            corrected_sample_values[crm_id].append(corrected_pivot)
                            lower_bounds[crm_id].append(lower)
                            upper_bounds[crm_id].append(upper)
                            soln_concs[crm_id].append(soln_conc)
                            int_values[crm_id].append(int_val)

            if not unique_crm_ids:
                self.main_plot.clear()
                self.legend.clear()
                self.logger.warning(f"No valid Verification data for {self.selected_element}")
                QMessageBox.warning(self, "Warning", f"No valid Verification data for {self.selected_element}")
                return

            self.main_plot.setLabel('bottom', 'Verification ID')
            self.main_plot.setLabel('left', f'{self.selected_element} Value')
            self.main_plot.setTitle(f'Verification Values for {self.selected_element}')
            self.main_plot.getAxis('bottom').setTicks([[(i, f'V {id}') for i, id in enumerate(unique_crm_ids)]])
            all_y_values = []
            for crm_id in unique_crm_ids:
                all_y_values.extend(certificate_values.get(crm_id, []))
                all_y_values.extend(sample_values.get(crm_id, []))
                all_y_values.extend(corrected_sample_values.get(crm_id, []))
                all_y_values.extend(lower_bounds.get(crm_id, []))
                all_y_values.extend(upper_bounds.get(crm_id, []))
            if all_y_values:
                y_min, y_max = min(all_y_values), max(all_y_values)
                margin = (y_max - y_min) * 0.1
                self.main_plot.setXRange(-0.5, len(unique_crm_ids) - 0.5)
                self.main_plot.setYRange(y_min - margin, y_max + margin)
                self.initial_ranges['main_x'] = (-0.5, len(unique_crm_ids) - 0.5)
                self.initial_ranges['main_y'] = (y_min - margin, y_max + margin)

            if self.show_check_crm.isChecked():
                for crm_id in unique_crm_ids:
                    x_pos = x_pos_map[crm_id]
                    cert_vals = certificate_values.get(crm_id, [])
                    if cert_vals:
                        x_vals = [x_pos] * len(cert_vals)
                        scatter = pg.PlotDataItem(
                            x=x_vals, y=cert_vals, pen=None, symbol='o', symbolSize=8,
                            symbolPen='g', symbolBrush='g'
                        )
                        self.main_plot.addItem(scatter)
                if 'Certificate Value' not in added_legend_names:
                    scatter = pg.PlotDataItem(
                        x=[-10], y=[0], pen=None, symbol='o', symbolSize=8,
                        symbolPen='g', symbolBrush='g', name='Certificate Value'
                    )
                    self.legend.addItem(scatter, 'Certificate Value')
                    added_legend_names.add('Certificate Value')

            if self.show_pivot_crm.isChecked():
                for crm_id in unique_crm_ids:
                    x_pos = x_pos_map[crm_id]
                    samp_vals = sample_values.get(crm_id, [])
                    if samp_vals:
                        x_vals = [x_pos] * len(samp_vals)
                        scatter = pg.PlotDataItem(
                            x=x_vals, y=samp_vals, pen=None, symbol='t', symbolSize=8,
                            symbolPen='b', symbolBrush='b'
                        )
                        self.main_plot.addItem(scatter)
                if 'Sample Value' not in added_legend_names:
                    scatter = pg.PlotDataItem(
                        x=[-10], y=[0], pen=None, symbol='t', symbolSize=8,
                        symbolPen='b', symbolBrush='b', name='Sample Value'
                    )
                    self.legend.addItem(scatter, 'Sample Value')
                    added_legend_names.add('Sample Value')

            if self.show_range.isChecked():
                for crm_id in unique_crm_ids:
                    x_pos = x_pos_map[crm_id]
                    low_bounds = lower_bounds.get(crm_id, [])
                    up_bounds = upper_bounds.get(crm_id, [])
                    if low_bounds and up_bounds:
                        for low, up in zip(low_bounds, up_bounds):
                            line_lower = pg.PlotDataItem(
                                x=[x_pos - 0.2, x_pos + 0.2], y=[low, low],
                                pen=pg.mkPen('r', width=2)
                            )
                            line_upper = pg.PlotDataItem(
                                x=[x_pos - 0.2, x_pos + 0.2], y=[up, up],
                                pen=pg.mkPen('r', width=2)
                            )
                            self.main_plot.addItem(line_lower)
                            self.main_plot.addItem(line_upper)
                if 'Acceptable Range' not in added_legend_names:
                    range_item = pg.PlotDataItem(
                        x=[-10 - 0.2, -10 + 0.2], y=[0, 0],
                        pen=pg.mkPen('r', width=2), name='Acceptable Range'
                    )
                    self.legend.addItem(range_item, 'Acceptable Range')
                    added_legend_names.add('Acceptable Range')

            self.main_plot.showGrid(x=True, y=True, alpha=0.3)

        except Exception as e:
            self.main_plot.clear()
            self.legend.clear()
            self.logger.error(f"Failed to update plot: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to update plot: {str(e)}")

    def open_select_crms_window(self):
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