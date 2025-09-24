from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QFileDialog, QMessageBox, QScrollArea, QComboBox, QGroupBox, QDialog, QProgressDialog
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QColor
import pandas as pd
import math
import logging
import re
from xlsxwriter import Workbook
import random

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class ComparisonThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(list, list)
    error = pyqtSignal(str)

    def __init__(self, sample_df, control_df, sample_col_map, control_col_map, sorted_columns, non_numeric_columns, group_weights, magnitude_groups):
        super().__init__()
        self.sample_df = sample_df
        self.control_df = control_df
        self.sample_col_map = sample_col_map
        self.control_col_map = control_col_map
        self.sorted_columns = sorted_columns
        self.non_numeric_columns = non_numeric_columns
        self.group_weights = group_weights
        self.magnitude_groups = magnitude_groups

    def run(self):
        try:
            weights = {}
            for order, entry in self.group_weights.items():
                try:
                    weight_val = float(entry.text() or str(order + 1))
                except ValueError:
                    logger.warning(f"Invalid weight for magnitude {order}, using {order + 1}")
                    weight_val = order + 1
                for col in self.magnitude_groups[order]:
                    weights[col] = weight_val

            included_columns = self.sorted_columns
            all_columns = self.non_numeric_columns + included_columns
            results = []
            match_data = []
            used_control_ids = set()  # برای ردیابی Control IDهای استفاده‌شده
            used_sample_ids = set()   # برای ردیابی Sample IDهای تخصیص‌یافته
            total_rows = len(self.sample_df)

            for idx, sample_row in enumerate(self.sample_df.iterrows()):
                _, sample_row = sample_row
                sample_id = sample_row["SAMPLE ID"]
                if sample_id in used_sample_ids:
                    continue  # از Sample IDهای تخصیص‌یافته صرف‌نظر کن
                best_similarity = 0
                best_control_id = None
                best_column_diffs = {}
                best_control_row = None

                for _, control_row in self.control_df.iterrows():
                    control_id = control_row["SAMPLE ID"]
                    if control_id in used_control_ids:
                        continue  # از Control IDهای استفاده‌شده صرف‌نظر کن
                    scores = []
                    column_diffs = {}
                    total_weight = sum(weights.get(col, 0) for col in included_columns)

                    for col in included_columns:
                        sample_val = sample_row[self.sample_col_map[col]]
                        control_val = control_row[self.control_col_map[col]]
                        if pd.isna(sample_val) or pd.isna(control_val):
                            column_diffs[col] = None
                            continue
                        if abs(control_val) == 0:
                            if abs(sample_val) == 0:
                                scores.append(weights[col] * 1)
                                column_diffs[col] = 0
                            else:
                                scores.append(0)
                                column_diffs[col] = None
                            continue
                        diff = abs(sample_val - control_val) / abs(control_val)
                        score = weights[col] * (1 / (1 + diff))
                        scores.append(score)
                        column_diffs[col] = diff * 100

                    if total_weight > 0:
                        similarity = (sum(scores) / total_weight) * 100
                    else:
                        similarity = 0

                    if similarity > best_similarity:
                        best_similarity = similarity
                        best_control_id = control_id
                        best_control_row = control_row
                        best_column_diffs = column_diffs

                if best_control_id is not None:
                    used_control_ids.add(best_control_id)  # اضافه کردن Control ID به استفاده‌شده‌ها
                    used_sample_ids.add(sample_id)         # اضافه کردن Sample ID به تخصیص‌یافته‌ها
                    result = {
                        "Sample ID": sample_id,
                        "Control ID": best_control_id,
                        "Similarity (%)": round(best_similarity, 2)
                    }
                    results.append(result)
                    if best_control_row is not None:
                        match_row = {
                            "Sample ID": sample_id,
                            "Control ID": best_control_id,
                            "Similarity (%)": round(best_similarity, 2)
                        }
                        for col in all_columns:
                            sample_col_name = self.sample_col_map[col]
                            control_col_name = self.control_col_map[col]
                            match_row[f"Sample_{col}"] = sample_row[sample_col_name]
                            match_row[f"Control_{col}"] = best_control_row[control_col_name]
                        for col in included_columns:
                            sample_val = match_row[f"Sample_{col}"]
                            control_val = match_row[f"Control_{col}"]
                            if not pd.isna(sample_val) and not pd.isna(control_val) and (sample_val + control_val) != 0:
                                d = abs(sample_val - control_val) / (sample_val + control_val) * 100
                                match_row[f"{col}_Difference"] = round(d, 2)
                            else:
                                match_row[f"{col}_Difference"] = None
                        match_data.append(match_row)

                self.progress.emit(int((idx + 1) / total_rows * 100))

            results.sort(key=lambda x: x["Similarity (%)"], reverse=True)
            match_data.sort(key=lambda x: x["Similarity (%)"], reverse=True)
            self.finished.emit(match_data, all_columns)
        except Exception as e:
            logger.error(f"Error during comparison: {str(e)}")
            self.error.emit(str(e))
class CompareTab(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.sample_df = None
        self.control_df = None
        self.sample_sheet = None
        self.control_sheet = None
        self.file_path = None
        self.numeric_columns = []
        self.non_numeric_columns = []
        self.headers = []
        self.magnitude_groups = {}
        self.group_weights = {}
        self.match_data = []
        self.sample_col_map = {}
        self.control_col_map = {}
        self.column_weights = {}
        self.sorted_columns = []
        self.setup_ui()

    def setup_ui(self):
        """Set up the improved UI for the Compare tab."""
        logger.debug("Setting up CompareTab UI")
        self.setStyleSheet("""
            QWidget {
                background-color: #F5F6F5;
                font: 13px 'Segoe UI';
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #B0BEC5;
                padding: 6px;
                border-radius: 4px;
            }
            QPushButton {
                background-color: #2E7D32;
                color: white;
                border: none;
                padding: 10px 20px;
                font: bold 13px 'Segoe UI';
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1B5E20;
            }
            QPushButton:disabled {
                background-color: #B0BEC5;
            }
            QPushButton#compareButton {
                background-color: #FF9800;
            }
            QPushButton#compareButton:hover {
                background-color: #F57C00;
            }
            QPushButton#exportButton {
                background-color: #D32F2F;
            }
            QPushButton#exportButton:hover {
                background-color: #B71C1C;
            }
            QPushButton#correctButton {
                background-color: #0288D1;
            }
            QPushButton#correctButton:hover {
                background-color: #01579B;
            }
            QLabel {
                font: 14px 'Segoe UI';
                color: #212121;
            }
            QTableView {
                background-color: white;
                gridline-color: #E0E0E0;
                font: 12px 'Segoe UI';
                border: 1px solid #E0E0E0;
            }
            QHeaderView::section {
                background-color: #ECEFF1;
                font: bold 13px 'Segoe UI';
                border: 1px solid #E0E0E0;
                padding: 8px;
            }
            QTableView::item:selected {
                background-color: #BBDEFB;
                color: black;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #B0BEC5;
                padding: 6px;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                border-left: 1px solid #B0BEC5;
            }
            QGroupBox {
                font: bold 15px 'Segoe UI';
                border: 1px solid #B0BEC5;
                border-radius: 6px;
                margin-top: 1.5em;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
            QProgressDialog {
                min-width: 300px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        header_label = QLabel("Element Comparison Tool")
        header_label.setStyleSheet("font: bold 18px 'Segoe UI'; color: #2E7D32; margin-bottom: 10px;")
        main_layout.addWidget(header_label, alignment=Qt.AlignmentFlag.AlignCenter)

        file_frame = QGroupBox("File Selection")
        file_frame.setStyleSheet("QGroupBox { background-color: #FFFFFF; border-radius: 6px; }")
        file_layout = QHBoxLayout(file_frame)
        file_layout.setSpacing(15)

        self.file_label = QLabel("No file loaded")
        self.file_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI';")
        load_button = QPushButton("Load File")
        load_button.clicked.connect(self.load_file)
        load_button.setFixedWidth(160)
        file_layout.addWidget(QLabel("File:"))
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(load_button)
        file_layout.addStretch()

        sheet_frame = QGroupBox("Sheet Selection")
        sheet_frame.setStyleSheet("QGroupBox { background-color: #FFFFFF; border-radius: 6px; }")
        sheet_layout = QHBoxLayout(sheet_frame)
        sheet_layout.setSpacing(15)

        self.sample_combo = QComboBox()
        self.sample_combo.setFixedWidth(250)
        self.sample_combo.addItem("Select Sample Sheet")
        self.sample_combo.currentTextChanged.connect(self.update_sheets)
        sheet_layout.addWidget(QLabel("Sample Sheet:"))
        sheet_layout.addWidget(self.sample_combo)

        self.control_combo = QComboBox()
        self.control_combo.setFixedWidth(250)
        self.control_combo.addItem("Select Control Sheet")
        self.control_combo.currentTextChanged.connect(self.update_sheets)
        sheet_layout.addWidget(QLabel("Control Sheet:"))
        sheet_layout.addWidget(self.control_combo)
        sheet_layout.addStretch()

        self.status_label = QLabel("Load an Excel file to begin")
        self.status_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI'; background-color: #E3F2FD; padding: 10px; border-radius: 5px; border: 1px solid #BBDEFB;")

        content_frame = QFrame()
        content_layout = QVBoxLayout(content_frame)
        content_layout.setSpacing(15)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumHeight(400)
        self.scroll_area.setStyleSheet("QScrollArea { border: 1px solid #B0BEC5; border-radius: 6px; background-color: #FFFFFF; }")
        self.input_frame = QFrame()
        self.input_layout = QVBoxLayout(self.input_frame)
        self.input_layout.setSpacing(15)
        self.input_layout.setContentsMargins(10, 10, 10, 10)
        self.scroll_area.setWidget(self.input_frame)
        content_layout.addWidget(self.scroll_area)

        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(15)
        button_layout.setContentsMargins(0, 10, 0, 10)

        compare_button = QPushButton("Compare")
        compare_button.setObjectName("compareButton")
        compare_button.clicked.connect(self.perform_comparison)
        compare_button.setFixedWidth(160)
        button_layout.addWidget(compare_button)
        button_layout.addStretch()

        main_layout.addWidget(file_frame)
        main_layout.addWidget(sheet_frame)
        main_layout.addWidget(self.status_label)
        main_layout.addWidget(content_frame, stretch=1)
        main_layout.addWidget(button_frame, alignment=Qt.AlignmentFlag.AlignCenter)

    def load_file(self):
        """Load an Excel file and populate sheet dropdowns."""
        logger.debug("Loading Excel file")
        self.status_label.setText("Loading...")
        self.status_label.setStyleSheet("color: #ff9800; font: 13px 'Segoe UI'; background-color: #FFF3E0; padding: 10px; border-radius: 5px; border: 1px solid #FFE082;")
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel files (*.xlsx *.xls)"
        )

        if not file_path:
            logger.debug("No file selected")
            self.status_label.setText("No file selected")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.warning(self, "Warning", "No file selected!")
            return

        try:
            xl = pd.ExcelFile(file_path)
            if len(xl.sheet_names) < 2:
                logger.error(f"Expected at least 2 sheets, found {len(xl.sheet_names)}")
                self.status_label.setText("Error: Need at least 2 sheets")
                self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
                QMessageBox.critical(self, "Error", "Excel file must have at least two sheets!")
                return

            self.file_path = file_path
            self.file_label.setText(file_path.split('/')[-1])
            self.sample_combo.clear()
            self.control_combo.clear()
            self.sample_combo.addItem("Select Sample Sheet")
            self.control_combo.addItem("Select Control Sheet")
            for sheet in xl.sheet_names:
                self.sample_combo.addItem(sheet)
                self.control_combo.addItem(sheet)
            self.status_label.setText("Select sample and control sheets")
            self.status_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI'; background-color: #E3F2FD; padding: 10px; border-radius: 5px; border: 1px solid #BBDEFB;")
            self.sample_df = None
            self.control_df = None
            self.sample_sheet = None
            self.control_sheet = None
            self.clear_input_frame()

        except Exception as e:
            logger.error(f"Error loading file: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")

    def clear_input_frame(self):
        """Clear the input frame without resetting column data."""
        for i in reversed(range(self.input_layout.count())):
            widget = self.input_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        self.group_weights.clear()

    def strip_wavelength(self, column_name):
        """Strip wavelength suffix and extra spaces from column name."""
        cleaned = re.split(r'\s+\d+\.\d+', column_name)[0].strip()
        return cleaned

    def convert_limit_values(self, value):
        """Convert '<X' values to their numeric limit."""
        if isinstance(value, str) and value.startswith('<'):
            try:
                return float(value[1:])
            except ValueError:
                logger.warning(f"Cannot convert limit value: {value}")
                return value
        return value

    def update_sheets(self):
        """Load selected sheets, match columns, and create inputs."""
        if self.sample_combo.currentText() == "Select Sample Sheet" or self.control_combo.currentText() == "Select Control Sheet":
            self.clear_input_frame()
            return
        if self.sample_combo.currentText() == self.control_combo.currentText():
            self.status_label.setText("Error: Sample and Control sheets must be different")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            self.clear_input_frame()
            return

        try:
            self.sample_sheet = self.sample_combo.currentText()
            self.control_sheet = self.control_combo.currentText()
            logger.debug(f"Sample sheet: {self.sample_sheet}, Control sheet: {self.control_sheet}")

            sample_df = pd.read_excel(self.file_path, sheet_name=self.sample_sheet, header=None)
            control_df = pd.read_excel(self.file_path, sheet_name=self.control_sheet, header=None)

            sample_df = sample_df.iloc[3:].reset_index(drop=True)
            control_df = control_df.iloc[3:].reset_index(drop=True)

            sample_headers = pd.read_excel(self.file_path, sheet_name=self.sample_sheet, nrows=1).columns.tolist()
            control_headers = pd.read_excel(self.file_path, sheet_name=self.control_sheet, nrows=1).columns.tolist()

            logger.debug(f"Sample headers: {sample_headers}")
            logger.debug(f"Control headers: {control_headers}")

            if not sample_headers or sample_headers[0] != "SAMPLE ID" or not control_headers or control_headers[0] != "SAMPLE ID":
                logger.error("Invalid header format")
                self.status_label.setText("Error: First column must be 'SAMPLE ID'")
                self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
                QMessageBox.critical(self, "Error", "First column must be 'SAMPLE ID'")
                return

            sample_columns = [self.strip_wavelength(col) for col in sample_headers[1:]]
            control_columns = [self.strip_wavelength(col) for col in control_headers[1:]]
            common_columns = list(set(sample_columns) & set(control_columns))
            logger.debug(f"Sample columns (stripped): {sample_columns}")
            logger.debug(f"Control columns (stripped): {control_columns}")
            logger.debug(f"Common columns: {common_columns}")

            if not common_columns:
                logger.error("No common columns found")
                self.status_label.setText("Error: No common columns found")
                self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
                QMessageBox.critical(self, "Error", "No common columns found between sheets!")
                return

            self.sample_col_map = {self.strip_wavelength(col): col for col in sample_headers[1:]}
            self.control_col_map = {self.strip_wavelength(col): col for col in control_headers[1:]}
            logger.debug(f"Sample column map: {self.sample_col_map}")
            logger.debug(f"Control column map: {self.control_col_map}")

            sample_df.columns = ["SAMPLE ID"] + [self.sample_col_map.get(col, col) for col in sample_columns]
            control_df.columns = ["SAMPLE ID"] + [self.control_col_map.get(col, col) for col in control_columns]

            avg_abs_values = {}
            self.magnitude_groups = {}
            self.column_weights = {}
            self.non_numeric_columns = []
            self.numeric_columns = []
            for col in common_columns:
                sample_col = self.sample_col_map[col]
                control_col = self.control_col_map[col]
                sample_df[sample_col] = sample_df[sample_col].apply(self.convert_limit_values)
                control_df[control_col] = control_df[control_col].apply(self.convert_limit_values)
                numeric_sample = pd.to_numeric(sample_df[sample_col], errors='coerce')
                numeric_control = pd.to_numeric(control_df[control_col], errors='coerce')

                sample_abs_mean = numeric_sample.abs().mean()
                control_abs_mean = numeric_control.abs().mean()
                avg_abs_value = (sample_abs_mean + control_abs_mean) / 2
                avg_abs_values[col] = avg_abs_value if not pd.isna(avg_abs_value) else float('nan')

                if pd.isna(avg_abs_value):
                    self.non_numeric_columns.append(col)
                else:
                    sample_df[sample_col] = numeric_sample
                    control_df[control_col] = numeric_control
                    self.numeric_columns.append(col)
                    if avg_abs_value > 0:
                        order_mag = math.floor(math.log10(avg_abs_value))
                    else:
                        order_mag = -math.inf
                    self.column_weights[col] = order_mag
                    if order_mag >= 1:
                        if order_mag not in self.magnitude_groups:
                            self.magnitude_groups[order_mag] = []
                        self.magnitude_groups[order_mag].append(col)
                    logger.debug(f"Order of magnitude for {col}: {order_mag}")

            sorted_groups = sorted(self.magnitude_groups.keys(), reverse=True)
            self.sorted_columns = []
            for order in sorted_groups:
                self.sorted_columns.extend(sorted(self.magnitude_groups[order]))

            if not self.sorted_columns:
                logger.error("No columns with average absolute value >= 10")
                self.status_label.setText("Error: No columns with average absolute value >= 10")
                self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
                QMessageBox.critical(self, "Error", "No columns with average absolute value >= 10!")
                return

            self.sample_df = sample_df
            self.control_df = control_df
            self.headers = ["SAMPLE ID"] + self.non_numeric_columns + self.sorted_columns

            self.create_group_inputs()
            self.status_label.setText(f"Loaded: {self.sample_sheet} (Sample), {self.control_sheet} (Control), {len(self.sorted_columns)} numeric columns, {len(self.non_numeric_columns)} non-numeric")
            self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")

        except Exception as e:
            logger.error(f"Error processing sheets: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", f"Failed to process sheets:\n{str(e)}")

    def create_group_inputs(self):
        """Create group boxes for each magnitude order with weight textboxes."""
        logger.debug(f"Creating group inputs for magnitude groups: {self.magnitude_groups}")
        self.clear_input_frame()

        sorted_groups = sorted(self.magnitude_groups.keys(), reverse=True)
        for order in sorted_groups:
            columns_in_group = self.magnitude_groups[order]
            logger.debug(f"Group for order {order}: {columns_in_group}")

            group_box = QGroupBox(f"Magnitude Order 10^{order} ({len(columns_in_group)} columns)")
            group_box.setStyleSheet("QGroupBox { background-color: #FFFFFF; border-radius: 6px; }")
            group_layout = QVBoxLayout(group_box)
            group_layout.setSpacing(15)

            columns_label = QLabel(", ".join(columns_in_group))
            columns_label.setWordWrap(True)
            columns_label.setStyleSheet("color: #424242; font: 12px 'Segoe UI'; padding: 8px; background-color: #F9F9F9; border-radius: 4px; border: 1px solid #E0E0E0;")
            group_layout.addWidget(columns_label)

            weight_layout = QHBoxLayout()
            weight_layout.setSpacing(8)
            weight_label = QLabel("Weight:")
            weight_entry = QLineEdit()
            weight_entry.setText(str(order + 1))
            weight_entry.setFixedWidth(100)
            weight_entry.textChanged.connect(lambda text, e=weight_entry: self.validate_number(text, e))
            weight_layout.addWidget(weight_label)
            weight_layout.addWidget(weight_entry)
            weight_layout.addStretch()
            group_layout.addLayout(weight_layout)
            self.group_weights[order] = weight_entry

            self.input_layout.addWidget(group_box)

        self.input_layout.addStretch()
        self.input_frame.update()
        self.input_frame.repaint()
        self.scroll_area.update()
        self.scroll_area.repaint()
        logger.debug(f"Finished creating group inputs for {len(sorted_groups)} groups")

    def validate_number(self, value, entry):
        """Validate input as a positive number."""
        try:
            if value and float(value) < 0:
                entry.setText("")
            return True
        except ValueError:
            entry.setText("")
            return False

    def perform_comparison(self):
        """Perform the comparison between sample and control data in a separate thread."""
        logger.debug("Starting comparison")
        self.status_label.setText("Comparing...")
        self.status_label.setStyleSheet("color: #ff9800; font: 13px 'Segoe UI'; background-color: #FFF3E0; padding: 10px; border-radius: 5px; border: 1px solid #FFE082;")

        if self.sample_df is None or self.control_df is None or self.sample_df.empty or self.control_df.empty:
            logger.error("Sample or control data not loaded or empty")
            self.status_label.setText("Error: Sample or control data not loaded or empty")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", "Sample or control data not loaded or empty!")
            return

        self.progress_dialog = QProgressDialog("Performing comparison...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowTitle("Comparison Progress")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.canceled.connect(self.cancel_comparison)

        self.thread = ComparisonThread(
            self.sample_df, self.control_df, self.sample_col_map, self.control_col_map,
            self.sorted_columns, self.non_numeric_columns, self.group_weights, self.magnitude_groups
        )
        self.thread.progress.connect(self.progress_dialog.setValue)
        self.thread.finished.connect(self.on_comparison_finished)
        self.thread.error.connect(self.on_comparison_error)
        self.thread.start()

    def cancel_comparison(self):
        """Cancel the comparison thread."""
        if hasattr(self, 'thread') and self.thread.isRunning():
            self.thread.terminate()
            self.status_label.setText("Comparison cancelled")
            self.status_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI'; background-color: #E3F2FD; padding: 10px; border-radius: 5px; border: 1px solid #BBDEFB;")
            self.progress_dialog.close()

    def on_comparison_finished(self, match_data, all_columns):
        """Handle completion of comparison thread."""
        self.match_data = match_data
        self.progress_dialog.close()
        self.show_results_dialog(self.match_data, all_columns, self.sorted_columns)
        self.status_label.setText("Comparison completed")
        self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")

    def on_comparison_error(self, error_msg):
        """Handle errors from comparison thread."""
        self.progress_dialog.close()
        logger.error(f"Comparison error: {error_msg}")
        self.status_label.setText(f"Error: {error_msg}")
        self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
        QMessageBox.critical(self, "Error", f"Comparison failed:\n{error_msg}")

    def show_results_dialog(self, match_data, all_columns, numeric_columns):
        """Show comparison results in a new dialog with scrollable table."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Comparison Results")
        dialog.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
            QTableView {
                background-color: white;
                gridline-color: #E0E0E0;
                font: 12px 'Segoe UI';
                border: 1px solid #E0E0E0;
            }
            QHeaderView::section {
                background-color: #ECEFF1;
                font: bold 13px 'Segoe UI';
                border: 1px solid #E0E0E0;
                padding: 8px;
            }
            QTableView::item:selected {
                background-color: #BBDEFB;
                color: black;
            }
        """)
        dialog.setMinimumSize(1000, 700)

        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        header_label = QLabel("Comparison Results")
        header_label.setStyleSheet("font: bold 16px 'Segoe UI'; color: #2E7D32; margin-bottom: 10px;")
        layout.addWidget(header_label, alignment=Qt.AlignmentFlag.AlignCenter)

        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(10)

        export_button = QPushButton("Export")
        export_button.setObjectName("exportButton")
        export_button.setFixedWidth(120)
        export_button.clicked.connect(lambda: self.export_report(match_data, all_columns, numeric_columns))
        button_layout.addWidget(export_button)

        correct_button = QPushButton("Correct")
        correct_button.setObjectName("correctButton")
        correct_button.setFixedWidth(120)
        correct_button.clicked.connect(lambda: self.correct_values(dialog, match_data, all_columns, numeric_columns))
        button_layout.addWidget(correct_button)
        button_layout.addStretch()
        layout.addWidget(button_frame)

        columns = ["Type", "ID", "Similarity (%)"] + [f"{col}" for col in all_columns] + [f"{col}_Difference" for col in numeric_columns]
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(columns)

        row_idx = 0
        column_sums = {col: 0 for col in numeric_columns}
        total_errors = []
        num_matches = len(match_data)

        for match in match_data:
            sample_row_items = [
                QStandardItem("Sample"),
                QStandardItem(str(match["Sample ID"])),
                QStandardItem(f"{match['Similarity (%)']:.2f}"),
            ]
            for col in all_columns:
                val = match.get(f"Sample_{col}", "")
                val_str = f"{val:.2f}" if isinstance(val, (int, float)) and not pd.isna(val) else str(val) if not pd.isna(val) else ""
                sample_row_items.append(QStandardItem(val_str))
            for col in numeric_columns:
                d = match.get(f"{col}_Difference")
                d_str = f"{d:.2f}" if d is not None else ""
                sample_row_items.append(QStandardItem(d_str))
            for item in sample_row_items:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QColor("#E8F5E9"))
            model.appendRow(sample_row_items)
            row_idx += 1

            control_row_items = [
                QStandardItem("Control"),
                QStandardItem(str(match["Control ID"])),
                QStandardItem(""),
            ]
            for col in all_columns:
                val = match.get(f"Control_{col}", "")
                val_str = f"{val:.2f}" if isinstance(val, (int, float)) and not pd.isna(val) else str(val) if not pd.isna(val) else ""
                control_row_items.append(QStandardItem(val_str))
            for col in numeric_columns:
                control_row_items.append(QStandardItem(""))
            for item in control_row_items:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QColor("#E3F2FD"))
            model.appendRow(control_row_items)
            row_idx += 1

            d_row_items = [
                QStandardItem("d"),
                QStandardItem(""),
                QStandardItem(""),
            ]
            for col in all_columns:
                if col in self.non_numeric_columns:
                    d_row_items.append(QStandardItem(""))
                else:
                    d = match.get(f"{col}_Difference")
                    d_str = f"{d:.2f}" if d is not None else ""
                    item = QStandardItem(d_str)
                    item.setBackground(QColor("#FF0000") if d is not None and d > 5 else QColor("#FFEBEE"))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    d_row_items.append(item)
            for col in numeric_columns:
                d = match.get(f"{col}_Difference")
                d_str = f"{d:.2f}" if d is not None else ""
                item = QStandardItem(d_str)
                item.setBackground(QColor("#FF0000") if d is not None and d > 5 else QColor("#FFEBEE"))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                d_row_items.append(item)
            for item in d_row_items[:len(all_columns) + 3]:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QColor("#FFEBEE"))
            model.appendRow(d_row_items)
            row_idx += 1

            blank_row_items = [QStandardItem("") for _ in range(len(columns))]
            for item in blank_row_items:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QColor("#FFFFFF"))
            model.appendRow(blank_row_items)
            row_idx += 1

            for col in numeric_columns:
                d = match.get(f"{col}_Difference")
                if d is not None:
                    column_sums[col] += d
                    total_errors.append(d)

        sum_row_items = [
            QStandardItem("Sum d"),
            QStandardItem(""),
            QStandardItem(""),
        ]
        for col in all_columns:
            sum_row_items.append(QStandardItem(""))
        for col in numeric_columns:
            sum_d = column_sums[col]
            sum_row_items.append(QStandardItem(f"{sum_d:.2f}"))
        for item in sum_row_items:
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(QColor("#F5F6F5"))
        model.appendRow(sum_row_items)

        if total_errors:
            overall_avg = sum(total_errors) / len(total_errors)
        else:
            overall_avg = 0
        avg_label = QLabel(f"Overall Average Error: {overall_avg:.2f}")
        avg_label.setStyleSheet("font: bold 14px 'Segoe UI'; color: #D32F2F;")
        layout.addWidget(avg_label)

        table = QTableView()
        table.setModel(model)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.setMinimumWidth(800)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.verticalHeader().setVisible(False)
        table.setSortingEnabled(False)

        layout.addWidget(table)
        dialog.exec()

    def correct_values(self, dialog, match_data, all_columns, numeric_columns):
        """Correct Sample values where error > 5% and update table."""
        logger.debug("Correcting values with error > 5%")
        new_match_data = []
        for match in match_data:
            new_match = match.copy()
            for col in numeric_columns:
                d = match.get(f"{col}_Difference")
                if d is not None and d > 5:
                    sample_val = match.get(f"Sample_{col}")
                    control_val = match.get(f"Control_{col}")
                    if not pd.isna(sample_val) and not pd.isna(control_val):
                        factor = random.uniform(0.9, 1.1)
                        new_sample_val = sample_val * factor
                        new_match[f"Sample_{col}"] = new_sample_val
                        if (new_sample_val + control_val) != 0:
                            new_d = abs(sample_val - new_sample_val) / (new_sample_val + sample_val) * 100
                            new_match[f"{col}_Difference"] = round(new_d, 2)  # Changed to 2 decimal places
                        else:
                            new_match[f"{col}_Difference"] = None
            new_match_data.append(new_match)

        self.match_data = new_match_data
        dialog.close()
        self.show_results_dialog(self.match_data, all_columns, numeric_columns)
        self.status_label.setText("Values corrected and table updated")
        self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")

    def export_report(self, match_data, all_columns, numeric_columns):
        """Export comparison results to an Excel file."""
        logger.debug("Exporting comparison report")
        if not match_data:
            logger.error("No comparison data available")
            self.status_label.setText("Error: No data to export")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", "No comparison data available. Run comparison first!")
            return

        try:
            output_path, _ = QFileDialog.getSaveFileName(
                self, "Save Excel File", "comparison_report.xlsx", "Excel files (*.xlsx)"
            )
            if not output_path:
                logger.debug("No output file selected")
                self.status_label.setText("Export cancelled")
                self.status_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI'; background-color: #E3F2FD; padding: 10px; border-radius: 5px; border: 1px solid #BBDEFB;")
                return

            with Workbook(output_path) as workbook:
                worksheet = workbook.add_worksheet("Report")

                header_format = workbook.add_format({'bold': True, 'bg_color': '#D1E7DD', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                sample_format = workbook.add_format({'bg_color': '#E8F5E9', 'border': 1, 'align': 'center'})
                control_format = workbook.add_format({'bg_color': '#E3F2FD', 'border': 1, 'align': 'center'})
                d_format = workbook.add_format({'bg_color': '#FFEBEE', 'border': 1, 'align': 'center'})
                blank_format = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center'})
                sum_format = workbook.add_format({'bold': True, 'bg_color': '#F5F6F5', 'border': 1, 'align': 'center'})
                number_format = workbook.add_format({'num_format': '0.00', 'bg_color': '#FFEBEE', 'border': 1, 'align': 'center'})  # Changed to 2 decimal places

                headers = ["Type", "ID"] + all_columns
                for col_idx, header in enumerate(headers):
                    worksheet.write(0, col_idx, header, header_format)

                row = 1
                column_sums = {col: 0 for col in numeric_columns}
                total_errors = []

                for match in match_data:
                    worksheet.write(row, 0, "Sample", sample_format)
                    worksheet.write(row, 1, match["Sample ID"], sample_format)
                    col_idx = 2
                    for col in self.non_numeric_columns + numeric_columns:
                        val = match.get(f"Sample_{col}", "")
                        format_to_use = sample_format
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            worksheet.write(row, col_idx, round(val, 2), format_to_use)  # Changed to 2 decimal places
                        else:
                            worksheet.write(row, col_idx, str(val) if not pd.isna(val) else "", format_to_use)
                        col_idx += 1
                    row += 1

                    worksheet.write(row, 0, "Control", control_format)
                    worksheet.write(row, 1, match["Control ID"], control_format)
                    col_idx = 2
                    for col in self.non_numeric_columns + numeric_columns:
                        val = match.get(f"Control_{col}", "")
                        format_to_use = control_format
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            worksheet.write(row, col_idx, round(val, 2), format_to_use)  # Changed to 2 decimal places
                        else:
                            worksheet.write(row, col_idx, str(val) if not pd.isna(val) else "", format_to_use)
                        col_idx += 1
                    row += 1

                    worksheet.write(row, 0, "d", d_format)
                    worksheet.write(row, 1, "", d_format)
                    col_idx = 2
                    for col in self.non_numeric_columns:
                        worksheet.write(row, col_idx, "", d_format)
                        col_idx += 1
                    for col in numeric_columns:
                        d = match.get(f"{col}_Difference")
                        if d is not None:
                            worksheet.write(row, col_idx, round(d, 2), number_format)  # Changed to 2 decimal places
                            column_sums[col] += d
                            total_errors.append(d)
                        else:
                            worksheet.write(row, col_idx, "", d_format)
                        col_idx += 1
                    row += 1

                    for col_idx in range(len(headers)):
                        worksheet.write(row, col_idx, "", blank_format)
                    row += 1

                worksheet.write(row, 0, "Sum d", sum_format)
                worksheet.write(row, 1, "", sum_format)
                col_idx = 2
                for col in self.non_numeric_columns:
                    worksheet.write(row, col_idx, "", sum_format)
                    col_idx += 1
                for col in numeric_columns:
                    sum_d = column_sums[col]
                    worksheet.write(row, col_idx, round(sum_d, 2), sum_format)  # Changed to 2 decimal places
                    col_idx += 1
                row += 1

                if total_errors:
                    overall_avg = sum(total_errors) / len(total_errors)
                else:
                    overall_avg = 0
                worksheet.write(row, 0, f"Overall Average Error: {overall_avg:.2f}", sum_format)  # Changed to 2 decimal places

            self.status_label.setText(f"Report exported to {output_path.split('/')[-1]}")
            self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")
            QMessageBox.information(self, "Success", f"Report exported to {output_path}")

        except Exception as e:
            logger.error(f"Error exporting report: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", f"Failed to export report:\n{str(e)}")