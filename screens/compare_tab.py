from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QFileDialog, QMessageBox, QScrollArea, QComboBox, QGroupBox, QDialog
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import pandas as pd
import math
import logging
import re
from xlsxwriter import Workbook
import uuid

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

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
        self.headers = []
        self.magnitude_groups = {}  # Groups of columns by magnitude order
        self.group_weights = {}  # Textboxes for weights
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
            QPushButton#exportButton {
                background-color: #D32F2F;
            }
            QPushButton#exportButton:hover {
                background-color: #B71C1C;
            }
            QPushButton#compareButton {
                background-color: #FF9800;
            }
            QPushButton#compareButton:hover {
                background-color: #F57C00;
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
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # Header label
        header_label = QLabel("Element Comparison Tool")
        header_label.setStyleSheet("font: bold 18px 'Segoe UI'; color: #2E7D32; margin-bottom: 10px;")
        main_layout.addWidget(header_label, alignment=Qt.AlignmentFlag.AlignCenter)

        # File selection frame
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

        # Sheet selection frame
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

        # Status bar
        self.status_label = QLabel("Load an Excel file to begin")
        self.status_label.setStyleSheet("color: #6c757d; font: 13px 'Segoe UI'; background-color: #E3F2FD; padding: 10px; border-radius: 5px; border: 1px solid #BBDEFB;")

        # Content frame
        content_frame = QFrame()
        content_layout = QVBoxLayout(content_frame)
        content_layout.setSpacing(15)

        # Scrollable input frame for magnitude groups
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

        # Button frame
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(15)
        button_layout.setContentsMargins(0, 10, 0, 10)

        compare_button = QPushButton("Compare")
        compare_button.setObjectName("compareButton")
        compare_button.clicked.connect(self.perform_comparison)
        compare_button.setFixedWidth(160)
        button_layout.addWidget(compare_button)

        export_button = QPushButton("Export")
        export_button.setObjectName("exportButton")
        export_button.clicked.connect(self.export_report)
        export_button.setFixedWidth(160)
        button_layout.addWidget(export_button)
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
        """Strip wavelength suffix and extra spaces from column name (e.g., 'Ag 328.068' -> 'Ag')."""
        cleaned = re.split(r'\s+\d+\.\d+', column_name)[0].strip()
        return cleaned

    def convert_limit_values(self, value):
        """Convert '<X' values to their numeric limit (e.g., '<2' -> 2)."""
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

            # Load sheets
            sample_df = pd.read_excel(self.file_path, sheet_name=self.sample_sheet, header=None)
            control_df = pd.read_excel(self.file_path, sheet_name=self.control_sheet, header=None)

            sample_df = sample_df.iloc[3:].reset_index(drop=True)
            control_df = control_df.iloc[3:].reset_index(drop=True)

            # Read headers
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

            # Strip wavelengths and extra spaces, find common columns
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

            # Map original column names to stripped names
            self.sample_col_map = {self.strip_wavelength(col): col for col in sample_headers[1:]}
            self.control_col_map = {self.strip_wavelength(col): col for col in control_headers[1:]}
            logger.debug(f"Sample column map: {self.sample_col_map}")
            logger.debug(f"Control column map: {self.control_col_map}")

            # Set column names
            sample_df.columns = ["SAMPLE ID"] + [self.sample_col_map.get(col, col) for col in sample_columns]
            control_df.columns = ["SAMPLE ID"] + [self.control_col_map.get(col, col) for col in control_columns]

            # Convert limit values and numeric columns, and calculate order of magnitude
            avg_abs_values = {}
            self.magnitude_groups = {}
            self.column_weights = {}
            for col in common_columns:
                sample_col = self.sample_col_map[col]
                control_col = self.control_col_map[col]
                sample_df[sample_col] = sample_df[sample_col].apply(self.convert_limit_values)
                control_df[control_col] = control_df[control_col].apply(self.convert_limit_values)
                sample_df[sample_col] = pd.to_numeric(sample_df[sample_col], errors='coerce')
                control_df[control_col] = pd.to_numeric(control_df[control_col], errors='coerce')

                # Calculate average absolute value for the column
                sample_abs_mean = sample_df[sample_col].abs().mean()
                control_abs_mean = control_df[control_col].abs().mean()
                avg_abs_value = (sample_abs_mean + control_abs_mean) / 2
                avg_abs_values[col] = avg_abs_value if not pd.isna(avg_abs_value) else 0
                logger.debug(f"Average absolute value for {col}: {avg_abs_value}")

                # Calculate order of magnitude
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

            # Sort groups by magnitude descending
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
            self.numeric_columns = self.sorted_columns
            self.headers = ["SAMPLE ID"] + self.sorted_columns

            self.create_group_inputs()
            self.status_label.setText(f"Loaded: {self.sample_sheet} (Sample), {self.control_sheet} (Control), {len(self.sorted_columns)} columns")
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

            # List of columns in group
            columns_label = QLabel(", ".join(columns_in_group))
            columns_label.setWordWrap(True)
            columns_label.setStyleSheet("color: #424242; font: 12px 'Segoe UI'; padding: 8px; background-color: #F9F9F9; border-radius: 4px; border: 1px solid #E0E0E0;")
            group_layout.addWidget(columns_label)

            # Weight frame
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
        """Perform the comparison between sample and control data using weighted difference formula."""
        logger.debug("Performing comparison")
        self.status_label.setText("Comparing...")
        self.status_label.setStyleSheet("color: #ff9800; font: 13px 'Segoe UI'; background-color: #FFF3E0; padding: 10px; border-radius: 5px; border: 1px solid #FFE082;")
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
            results = []
            self.match_data = []
            for _, sample_row in self.sample_df.iterrows():
                sample_id = sample_row["SAMPLE ID"]
                best_similarity = 0
                best_control_id = None
                best_control_row = None
                best_column_diffs = {}

                for _, control_row in self.control_df.iterrows():
                    control_id = control_row["SAMPLE ID"]
                    scores = []
                    column_diffs = {}
                    total_weight = sum(weights[col] for col in included_columns)

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
                        column_diffs[col] = diff * 100  # Store as percentage

                    if total_weight > 0:
                        similarity = (sum(scores) / total_weight) * 100
                    else:
                        similarity = 0

                    if similarity > best_similarity:
                        best_similarity = similarity
                        best_control_id = control_id
                        best_control_row = control_row
                        best_column_diffs = column_diffs

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
                    for col in included_columns:
                        sample_val = sample_row[self.sample_col_map[col]]
                        control_val = best_control_row[self.control_col_map[col]]
                        match_row[f"Sample_{col}"] = sample_val
                        match_row[f"Control_{col}"] = control_val
                        match_row[f"{col}_Difference"] = best_column_diffs.get(col)
                    self.match_data.append(match_row)

            # Sort results by similarity (descending)
            results.sort(key=lambda x: x["Similarity (%)"], reverse=True)
            self.match_data.sort(key=lambda x: x["Similarity (%)"], reverse=True)

            self.show_results_dialog(results, included_columns)
            self.status_label.setText("Comparison completed")
            self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")

        except Exception as e:
            logger.error(f"Error during comparison: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", f"Comparison failed:\n{str(e)}")

    def show_results_dialog(self, results, included_columns):
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

        # Header label
        header_label = QLabel("Comparison Results")
        header_label.setStyleSheet("font: bold 16px 'Segoe UI'; color: #2E7D32; margin-bottom: 10px;")
        layout.addWidget(header_label, alignment=Qt.AlignmentFlag.AlignCenter)

        columns = ["Sample ID", "Control ID", "Similarity (%)"]
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(columns)

        for idx, result in enumerate(results):
            row = [
                QStandardItem(str(result["Sample ID"])),
                QStandardItem(str(result["Control ID"] or "")),
                QStandardItem(f"{result['Similarity (%)']:.2f}"),
            ]
            for item in row:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                # Set background color for alternating rows
                item.setBackground(Qt.GlobalColor.lightGray if idx % 2 else Qt.GlobalColor.white)
            model.appendRow(row)

        table = QTableView()
        table.setModel(model)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.setMinimumWidth(800)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.verticalHeader().setVisible(False)
        table.setSortingEnabled(True)
        table.sortByColumn(2, Qt.SortOrder.DescendingOrder)

        layout.addWidget(table)
        dialog.exec()

    def export_report(self):
        """Export comparison results and matching data to an Excel file, sorted by similarity."""
        logger.debug("Exporting comparison report")
        if not self.match_data:
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

            included_columns = self.sorted_columns
            summary_data = []
            for result in self.match_data:
                row = {
                    "Sample ID": result["Sample ID"],
                    "Control ID": result["Control ID"],
                    "Similarity (%)": result["Similarity (%)"]
                }
                summary_data.append(row)
            summary_df = pd.DataFrame(summary_data)

            details_data = []
            for result in self.match_data:
                row = {
                    "Sample ID": result["Sample ID"],
                    "Control ID": result["Control ID"],
                    "Similarity (%)": result["Similarity (%)"]
                }
                for col in included_columns:
                    row[f"Sample_{col}"] = result.get(f"Sample_{col}")
                    row[f"Control_{col}"] = result.get(f"Control_{col}")
                    row[f"{col}_Difference"] = result.get(f"{col}_Difference")
                details_data.append(row)
            details_df = pd.DataFrame(details_data)

            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                details_df.to_excel(writer, sheet_name="Details", index=False)

                workbook = writer.book
                worksheet_summary = writer.sheets["Summary"]
                worksheet_details = writer.sheets["Details"]
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#D1E7DD', 'border': 1, 'align': 'center', 'valign': 'vcenter'
                })
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                number_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00'})

                for col_num, value in enumerate(summary_df.columns.values):
                    worksheet_summary.write(0, col_num, value, header_format)
                for col_num, value in enumerate(details_df.columns.values):
                    worksheet_details.write(0, col_num, value, header_format)

                worksheet_summary.set_column(0, len(summary_df.columns) - 1, 15, cell_format)
                worksheet_details.set_column(0, 0, 15, cell_format)
                worksheet_details.set_column(1, 1, 15, cell_format)
                worksheet_details.set_column(2, 2, 15, number_format)
                col_idx = 3
                for col in included_columns:
                    worksheet_details.set_column(col_idx, col_idx + 2, 15, number_format)
                    col_idx += 3

            self.status_label.setText(f"Report exported to {output_path.split('/')[-1]}")
            self.status_label.setStyleSheet("color: #2e7d32; font: 13px 'Segoe UI'; background-color: #E8F5E9; padding: 10px; border-radius: 5px; border: 1px solid #A5D6A7;")
            QMessageBox.information(self, "Success", f"Report exported to {output_path}")

        except Exception as e:
            logger.error(f"Error exporting report: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 13px 'Segoe UI'; background-color: #FFEBEE; padding: 10px; border-radius: 5px; border: 1px solid #EF9A9A;")
            QMessageBox.critical(self, "Error", f"Failed to export report:\n{str(e)}")