from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QFileDialog, QMessageBox, QScrollArea, QCheckBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import pandas as pd
import math
import logging
from xlsxwriter import Workbook

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class CompareTab(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.sample_df = None
        self.control_df = None
        self.numeric_columns = []
        self.headers = []
        self.column_ranges = {}
        self.column_checkboxes = {}
        self.match_data = []
        self.setup_ui()

    def setup_ui(self):
        """Set up the UI for the Compare tab."""
        logger.debug("Setting up CompareTab UI")
        self.setStyleSheet("""
            QWidget {
                background-color: #FAFAFA;
                font: 12px 'Segoe UI';
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #B0BEC5;
                padding: 5px;
                border-radius: 4px;
            }
            QPushButton {
                background-color: #2E7D32;
                color: white;
                border: none;
                padding: 8px 16px;
                font: bold 11px 'Segoe UI';
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1B5E20;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
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
                font: 13px 'Segoe UI';
                color: #212121;
            }
            QTableView {
                background-color: white;
                gridline-color: #ccc;
                font: 11px 'Segoe UI';
            }
            QHeaderView::section {
                background-color: #ECEFF1;
                font: bold 12px 'Segoe UI';
                border: 1px solid #ccc;
                padding: 5px;
            }
            QTableView::item:selected {
                background-color: #BBDEFB;
                color: black;
            }
            QCheckBox {
                font: 11px 'Segoe UI';
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Button frame
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(10)
        load_button = QPushButton("📂 Load File")
        load_button.clicked.connect(self.compare_files)
        load_button.setFixedWidth(120)
        button_layout.addWidget(load_button)
        button_layout.addStretch()
        export_button = QPushButton("📊 Export")
        export_button.setObjectName("exportButton")
        export_button.clicked.connect(self.export_report)
        export_button.setFixedWidth(120)
        button_layout.addWidget(export_button)
        main_layout.addWidget(button_frame)

        # Status bar
        self.status_label = QLabel("No file loaded")
        self.status_label.setStyleSheet("color: #6c757d; font: 11px 'Segoe UI';")
        main_layout.addWidget(self.status_label)

        # Content frame
        content_frame = QFrame()
        content_layout = QVBoxLayout(content_frame)
        content_layout.setSpacing(5)
        main_layout.addWidget(content_frame, stretch=1)

        # Scrollable input frame
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.input_frame = QFrame()
        self.input_layout = QVBoxLayout(self.input_frame)
        self.input_layout.setSpacing(5)
        scroll_area.setWidget(self.input_frame)
        content_layout.addWidget(scroll_area)

        # Result frame
        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        self.result_label = QLabel("Load an Excel file to compare sheets")
        self.result_label.setStyleSheet("color: #6c757d; font: 12px 'Segoe UI';")
        result_layout.addWidget(self.result_label)
        content_layout.addWidget(self.result_frame, stretch=1)

    def compare_files(self):
        """Load and compare two sheets from an Excel file."""
        logger.debug("Initiating file comparison")
        self.status_label.setText("Loading...")
        self.status_label.setStyleSheet("color: #ff9800; font: 11px 'Segoe UI';")
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel files (*.xlsx *.xls)"
        )

        if not file_path:
            logger.debug("No file selected")
            self.status_label.setText("No file selected")
            self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
            QMessageBox.warning(self, "Warning", "No file selected!")
            return

        try:
            xl = pd.ExcelFile(file_path)
            if len(xl.sheet_names) != 2:
                logger.error(f"Expected 2 sheets, found {len(xl.sheet_names)}")
                self.status_label.setText("Error: Need exactly 2 sheets")
                self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
                QMessageBox.critical(self, "Error", "Excel file must have exactly two sheets!")
                return

            sheet1 = pd.read_excel(file_path, sheet_name=0, header=None)
            sheet2 = pd.read_excel(file_path, sheet_name=1, header=None)

            if sheet1.shape[0] >= sheet2.shape[0]:
                sample_df, control_df = sheet1, sheet2
                sample_sheet_name, control_sheet_name = xl.sheet_names[0], xl.sheet_names[1]
            else:
                sample_df, control_df = sheet2, sheet1
                sample_sheet_name, control_sheet_name = xl.sheet_names[1], xl.sheet_names[0]

            logger.debug(f"Sample: {sample_sheet_name}, Control: {control_sheet_name}")

            sample_df = sample_df.iloc[3:].reset_index(drop=True)
            control_df = control_df.iloc[3:].reset_index(drop=True)

            headers = pd.read_excel(file_path, sheet_name=sample_sheet_name, nrows=1).columns.tolist()
            if not headers or headers[0] != "SAMPLE ID":
                logger.error("Invalid header format")
                self.status_label.setText("Error: First column must be 'SAMPLE ID'")
                self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
                QMessageBox.critical(self, "Error", "First column must be 'SAMPLE ID'")
                return

            sample_df.columns = headers
            control_df.columns = headers

            numeric_columns = headers[1:]
            for col in numeric_columns:
                sample_df[col] = pd.to_numeric(sample_df[col], errors='coerce')
                control_df[col] = pd.to_numeric(control_df[col], errors='coerce')

            self.sample_df = sample_df
            self.control_df = control_df
            self.numeric_columns = numeric_columns
            self.headers = headers

            self.create_range_inputs(numeric_columns)
            self.status_label.setText(f"Loaded: {sample_sheet_name} (Sample), {control_sheet_name} (Control)")
            self.status_label.setStyleSheet("color: #2e7d32; font: 11px 'Segoe UI';")
            self.perform_comparison()

        except Exception as e:
            logger.error(f"Error loading file: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
            QMessageBox.critical(self, "Error", f"Failed to process file:\n{str(e)}")

    def create_range_inputs(self, columns):
        """Create input fields for ranges and checkboxes in 10 columns."""
        logger.debug("Creating range input fields")
        for i in reversed(range(self.input_layout.count())):
            widget = self.input_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        max_columns_per_row = 10
        num_rows = math.ceil(len(columns) / max_columns_per_row)
        self.column_ranges = {}
        self.column_checkboxes = {}

        for row in range(num_rows):
            start_idx = row * max_columns_per_row
            end_idx = min(start_idx + max_columns_per_row, len(columns))
            row_columns = columns[start_idx:end_idx]

            row_frame = QFrame()
            row_layout = QHBoxLayout(row_frame)
            row_layout.setSpacing(10)
            for col in row_columns:
                col_frame = QFrame()
                col_layout = QVBoxLayout(col_frame)
                col_layout.setSpacing(2)
                col_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

                label = QLabel(col)
                label.setStyleSheet("font: bold 11px 'Segoe UI'; color: #343a40;")
                col_layout.addWidget(label)

                entry = QLineEdit()
                entry.setText("15")
                entry.setFixedWidth(80)
                entry.textChanged.connect(lambda text, e=entry: self.validate_number(text, e))
                col_layout.addWidget(entry)
                self.column_ranges[col] = entry

                checkbox = QCheckBox()
                checkbox.setChecked(True)
                col_layout.addWidget(checkbox)
                self.column_checkboxes[col] = checkbox

                row_layout.addWidget(col_frame)
            row_layout.addStretch()
            self.input_layout.addWidget(row_frame)

        compare_button = QPushButton("🔍 Compare")
        compare_button.setObjectName("compareButton")
        compare_button.clicked.connect(self.perform_comparison)
        compare_button.setFixedWidth(120)
        self.input_layout.addWidget(compare_button, alignment=Qt.AlignmentFlag.AlignCenter)
        self.input_layout.addStretch()

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
        """Perform the comparison between sample and control data."""
        logger.debug("Performing comparison")
        self.status_label.setText("Comparing...")
        self.status_label.setStyleSheet("color: #ff9800; font: 11px 'Segoe UI';")
        try:
            ranges = {}
            included_columns = []
            for col in self.numeric_columns:
                try:
                    range_val = float(self.column_ranges[col].text() or "15")
                    ranges[col] = range_val / 100
                    if self.column_checkboxes[col].isChecked():
                        included_columns.append(col)
                except ValueError:
                    logger.warning(f"Invalid range for {col}, using 15%")
                    ranges[col] = 0.15
                    if self.column_checkboxes[col].isChecked():
                        included_columns.append(col)

            if not included_columns:
                logger.error("No columns selected")
                self.status_label.setText("Error: Select at least one column")
                self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
                QMessageBox.critical(self, "Error", "At least one column must be included!")
                return

            results = []
            self.match_data = []
            for _, sample_row in self.sample_df.iterrows():
                sample_id = sample_row["SAMPLE ID"]
                best_similarity = 0
                best_control_id = None
                best_column_status = {}
                best_control_row = None

                for _, control_row in self.control_df.iterrows():
                    control_id = control_row["SAMPLE ID"]
                    matches = 0
                    total_valid = 0
                    column_status = {}

                    for col in included_columns:
                        sample_val = sample_row[col]
                        control_val = control_row[col]
                        if pd.isna(sample_val) or pd.isna(control_val):
                            column_status[col] = "✗"
                            continue
                        total_valid += 1
                        range_val = ranges[col] * control_val
                        if abs(sample_val - control_val) <= range_val:
                            matches += 1
                            column_status[col] = "✓"
                        else:
                            column_status[col] = "✗"

                    if total_valid > 0:
                        similarity = (matches / total_valid) * 100
                        if similarity > best_similarity:
                            best_similarity = similarity
                            best_control_id = control_id
                            best_column_status = column_status
                            best_control_row = control_row

                result = {
                    "Sample ID": sample_id,
                    "Control ID": best_control_id,
                    "Similarity (%)": round(best_similarity, 2)
                }
                result.update({f"{col}_Status": best_column_status.get(col, "✗") for col in included_columns})
                results.append(result)
                if best_control_row is not None:
                    match_row = {
                        "Sample ID": sample_id,
                        "Control ID": best_control_id,
                        "Similarity (%)": round(best_similarity, 2)
                    }
                    for col in included_columns:
                        sample_val = sample_row[col]
                        control_val = best_control_row[col]
                        match_row[f"Sample_{col}"] = sample_val
                        match_row[f"Control_{col}"] = control_val
                        match_row[f"{col}_Difference"] = (
                            ((sample_val - control_val) / sample_val * 100)
                            if pd.notna(sample_val) and pd.notna(control_val) and sample_val != 0 else None
                        )
                        match_row[f"{col}_Status"] = best_column_status.get(col, "✗")
                    self.match_data.append(match_row)

            self.display_results(results, included_columns)
            self.status_label.setText("Comparison completed")
            self.status_label.setStyleSheet("color: #2e7d32; font: 11px 'Segoe UI';")

        except Exception as e:
            logger.error(f"Error during comparison: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
            QMessageBox.critical(self, "Error", f"Comparison failed:\n{str(e)}")

    def display_results(self, results, included_columns):
        """Display comparison results in a table."""
        logger.debug("Displaying comparison results")
        for i in reversed(range(self.result_frame.layout().count())):
            widget = self.result_frame.layout().itemAt(i).widget()
            if widget:
                widget.deleteLater()

        columns = ["Sample ID", "Control ID", "Similarity (%)"] + [f"{col}_Status" for col in included_columns]
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(columns)

        for idx, result in enumerate(results):
            row = [
                QStandardItem(str(result["Sample ID"])),
                QStandardItem(str(result["Control ID"] or "")),
                QStandardItem(f"{result['Similarity (%)']:.2f}"),
            ]
            row.extend([QStandardItem(result[f"{col}_Status"]) for col in included_columns])
            for item in row:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            model.appendRow(row)
            for col_idx, item in enumerate(row):
                model.setBackground(model.index(idx, col_idx), Qt.GlobalColor.lightGray if idx % 2 else Qt.GlobalColor.white)

        table = QTableView()
        table.setModel(model)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setColumnWidth(0, 120)
        table.setColumnWidth(1, 120)
        table.setColumnWidth(2, 100)
        for col_idx in range(3, len(columns)):
            table.setColumnWidth(col_idx, 50)
        table.verticalHeader().setVisible(False)
        table.setSortingEnabled(True)
        self.result_frame.layout().addWidget(table)

    def export_report(self):
        """Export comparison results and matching data to an Excel file."""
        logger.debug("Exporting comparison report")
        if not self.match_data:
            logger.error("No comparison data available")
            self.status_label.setText("Error: No data to export")
            self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
            QMessageBox.critical(self, "Error", "No comparison data available. Run comparison first!")
            return

        try:
            output_path, _ = QFileDialog.getSaveFileName(
                self, "Save Excel File", "comparison_report.xlsx", "Excel files (*.xlsx)"
            )
            if not output_path:
                logger.debug("No output file selected")
                self.status_label.setText("Export cancelled")
                self.status_label.setStyleSheet("color: #6c757d; font: 11px 'Segoe UI';")
                return

            summary_data = []
            for result in self.match_data:
                row = {
                    "Sample ID": result["Sample ID"],
                    "Control ID": result["Control ID"],
                    "Similarity (%)": result["Similarity (%)"]
                }
                row.update({f"{col}_Status": result.get(f"{col}_Status", "✗") for col in self.numeric_columns if col in self.column_checkboxes})
                summary_data.append(row)
            summary_df = pd.DataFrame(summary_data)

            details_data = []
            for result in self.match_data:
                row = {
                    "Sample ID": result["Sample ID"],
                    "Control ID": result["Control ID"],
                    "Similarity (%)": result["Similarity (%)"]
                }
                for col in self.numeric_columns:
                    if col in self.column_checkboxes:
                        row[f"Sample_{col}"] = result.get(f"Sample_{col}")
                        row[f"Control_{col}"] = result.get(f"Control_{col}")
                        row[f"{col}_Difference"] = result.get(f"{col}_Difference")
                        row[f"{col}_Status"] = result.get(f"{col}_Status", "✗")
                details_data.append(row)
            details_df = pd.DataFrame(details_data)

            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                details_df.to_excel(writer, sheet_name="Details", index=False)

                workbook = writer.book
                worksheet_summary = writer.sheets["Summary"]
                worksheet_details = writer.sheets["Details"]
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#d1e7dd', 'border': 1, 'align': 'center', 'valign': 'vcenter'
                })
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                number_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00'})

                for col_num, value in enumerate(summary_df.columns.values):
                    worksheet_summary.write(0, col_num, value, header_format)
                for col_num, value in enumerate(details_df.columns.values):
                    worksheet_details.write(0, col_num, value, header_format)

                worksheet_summary.set_column(0, len(summary_df.columns) - 1, 12, cell_format)
                worksheet_details.set_column(0, 0, 12, cell_format)
                worksheet_details.set_column(1, 1, 12, cell_format)
                worksheet_details.set_column(2, 2, 12, number_format)
                col_idx = 3
                for col in self.numeric_columns:
                    if col in self.column_checkboxes:
                        worksheet_details.set_column(col_idx, col_idx + 2, 12, number_format)
                        worksheet_details.set_column(col_idx + 3, col_idx + 3, 12, cell_format)
                        col_idx += 4

            self.status_label.setText("Report exported")
            self.status_label.setStyleSheet("color: #2e7d32; font: 11px 'Segoe UI';")
            QMessageBox.information(self, "Success", f"Report exported to {output_path}")

        except Exception as e:
            logger.error(f"Error exporting report: {str(e)}")
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #d32f2f; font: 11px 'Segoe UI';")
            QMessageBox.critical(self, "Error", f"Failed to export report:\n{str(e)}")