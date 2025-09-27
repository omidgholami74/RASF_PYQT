from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QGroupBox, QMessageBox, QProgressDialog
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont
import pandas as pd
import numpy as np
import time
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class WeightCorrectionThread(QThread):
    """Thread for applying weight corrections in the background."""
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, int)
    error = pyqtSignal(str)

    def __init__(self, df, solution_labels, new_weight, apply_to_all=False):
        super().__init__()
        self.df = df.copy()
        self.solution_labels = solution_labels
        self.new_weight = new_weight
        self.apply_to_all = apply_to_all

    def run(self):
        try:
            corrected_rows = 0
            total_rows = len(self.solution_labels)
            for i, solution_label in enumerate(self.solution_labels):
                mask = (self.df['Solution Label'] == solution_label) & (self.df['Type'] == 'Samp')
                matching_rows = self.df[mask]
                if not matching_rows.empty:
                    for idx in matching_rows.index:
                        current_weight = float(self.df.loc[idx, 'Act Wgt'])
                        current_corr_con = float(self.df.loc[idx, 'Corr Con'])
                        corrected_corr_con = (self.new_weight / current_weight) * current_corr_con
                        self.df.loc[idx, 'Corr Con'] = corrected_corr_con
                        self.df.loc[idx, 'Act Wgt'] = self.new_weight
                    corrected_rows += len(matching_rows)
                if self.apply_to_all:
                    self.progress.emit(int((i + 1) / total_rows * 100))
            self.finished.emit(self.df.to_json(), corrected_rows)
        except Exception as e:
            self.error.emit(str(e))

class WeightCheckFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.df_cache = None
        self.bad_weights = None
        self.correction_weight = {}
        self.selected_solution_label = None
        self.weight_min = 0.190
        self.weight_max = 0.210
        self.new_weight = 0.2
        self.setup_ui()

    def setup_ui(self):
        """Set up the UI with enhanced controls and a modern layout."""
        start_time = time.time()
        self.setStyleSheet("""
            QWidget {
                background-color: #F5F7FA;
                font-family: 'Inter', 'Segoe UI', sans-serif;
                font-size: 13px;
            }
            QGroupBox {
                font-weight: bold;
                color: #1A3C34;
                margin-top: 15px;
                border: 1px solid #D0D7DE;
                border-radius: 6px;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                left: 10px;
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #D0D7DE;
                padding: 6px;
                border-radius: 6px;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 1px solid #2E7D32;
                box-shadow: 0 0 5px rgba(46, 125, 50, 0.3);
            }
            QPushButton {
                background-color: #2E7D32;
                color: white;
                border: none;
                padding: 8px 16px;
                font-weight: 600;
                font-size: 13px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1B5E20;
            }
            QPushButton:disabled {
                background-color: #E0E0E0;
                color: #6B7280;
            }
            QLabel {
                color: #1A3C34;
                font-size: 13px;
            }
            QTableView {
                background-color: #FFFFFF;
                border: 1px solid #D0D7DE;
                gridline-color: #E5E7EB;
                font-size: 12px;
                selection-background-color: #DBEAFE;
                selection-color: #1A3C34;
            }
            QHeaderView::section {
                background-color: #F9FAFB;
                font-weight: 600;
                color: #1A3C34;
                border: 1px solid #D0D7DE;
                padding: 6px;
            }
            QTableView::item:selected {
                background-color: #DBEAFE;
                color: #1A3C34;
            }
            QTableView::item {
                padding: 0px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # Input group
        input_group = QGroupBox("Weight Range Check")
        input_layout = QHBoxLayout(input_group)
        input_layout.setSpacing(10)

        input_layout.addWidget(QLabel("Min Weight:"))
        self.min_weight_entry = QLineEdit()
        self.min_weight_entry.setText(str(self.weight_min))
        self.min_weight_entry.setFixedWidth(120)
        self.min_weight_entry.setToolTip("Enter the minimum acceptable weight (e.g., 0.190)")
        input_layout.addWidget(self.min_weight_entry)

        input_layout.addWidget(QLabel("Max Weight:"))
        self.max_weight_entry = QLineEdit()
        self.max_weight_entry.setText(str(self.weight_max))
        self.max_weight_entry.setFixedWidth(120)
        self.max_weight_entry.setToolTip("Enter the maximum acceptable weight (e.g., 0.210)")
        input_layout.addWidget(self.max_weight_entry)

        check_button = QPushButton("Check Weights")
        check_button.setToolTip("Check weights against the specified range")
        check_button.clicked.connect(self.check_weights)
        input_layout.addWidget(check_button)
        input_layout.addStretch()

        main_layout.addWidget(input_group)

        # Main container
        main_container = QFrame()
        main_layout.addWidget(main_container, stretch=1)
        container_layout = QHBoxLayout(main_container)
        container_layout.setSpacing(15)

        # Correction group
        correction_group = QGroupBox("Weight Correction")
        correction_layout = QVBoxLayout(correction_group)
        correction_layout.setSpacing(10)

        # New weight input
        new_weight_frame = QFrame()
        new_weight_layout = QHBoxLayout(new_weight_frame)
        new_weight_layout.setSpacing(10)
        new_weight_layout.addWidget(QLabel("New Weight:"))
        self.new_weight_entry = QLineEdit()
        self.new_weight_entry.setText(str(self.new_weight))
        self.new_weight_entry.setFixedWidth(120)
        self.new_weight_entry.setToolTip("Enter the new weight to apply to the selected sample")
        new_weight_layout.addWidget(self.new_weight_entry)
        
        correction_button = QPushButton("Apply Correction")
        correction_button.setToolTip("Apply the new weight to the selected sample")
        correction_button.clicked.connect(self.apply_weight_correction)
        new_weight_layout.addWidget(correction_button)

        apply_all_button = QPushButton("Apply to All")
        apply_all_button.setToolTip("Apply the new weight to all non-excluded samples in the table")
        apply_all_button.clicked.connect(self.apply_to_all)
        new_weight_layout.addWidget(apply_all_button)
        
        new_weight_layout.addStretch()
        correction_layout.addWidget(new_weight_frame)

        # Bad weights table
        self.correction_table = QTableView()
        self.correction_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.correction_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.correction_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.correction_table.verticalHeader().setVisible(False)
        self.correction_table.clicked.connect(self.select_row)
        self.correction_table.setToolTip("Select a row to correct its weight or mark it for exclusion")
        correction_layout.addWidget(self.correction_table)

        container_layout.addWidget(correction_group, stretch=1)

        logger.debug(f"UI setup took {time.time() - start_time:.3f} seconds")

    def select_row(self, index):
        """Handle row selection in the correction table."""
        if not index.isValid():
            return
        model = self.correction_table.model()
        self.selected_solution_label = model.data(model.index(index.row(), 1))
        if self.bad_weights is not None:
            weight_row = self.bad_weights[self.bad_weights['Solution Label'] == self.selected_solution_label]
            if not weight_row.empty:
                try:
                    actual_weight = float(weight_row['Act Wgt'].iloc[0])
                    self.new_weight_entry.setText(f"{actual_weight:.3f}")
                except (ValueError, TypeError) as e:
                    logger.error(f"Error setting new weight: {e}")
                    self.new_weight_entry.setText("0.2")

    def check_weights(self):
        """Check weights and display bad weights in the table."""
        start_time = time.time()
        try:
            self.weight_min = float(self.min_weight_entry.text())
            self.weight_max = float(self.max_weight_entry.text())
            if self.weight_min >= self.weight_max:
                raise ValueError("Minimum weight must be less than maximum weight")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid weight range: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading took {time.time() - data_start:.3f} seconds")

        df = self.df_cache
        if df is None or df.empty:
            QMessageBox.warning(self, "Warning", "No data loaded!")
            return

        # Convert 'Corr Con' to numeric and drop non-numeric rows
        df['Corr Con'] = pd.to_numeric(df['Corr Con'], errors='coerce')
        self.df_cache = df[df['Corr Con'].notna()].copy()
        df = self.df_cache

        # Filter bad weights
        data_filter_start = time.time()
        sample_data = df[df['Type'] == 'Samp']
        self.bad_weights = sample_data[
            (sample_data['Act Wgt'] < self.weight_min) | (sample_data['Act Wgt'] > self.weight_max)
        ][['Solution Label', 'Act Wgt', 'Corr Con']].drop_duplicates(subset=['Solution Label'])
        logger.debug(f"Filtering bad weights took {time.time() - data_filter_start:.3f} seconds")

        # Update table
        self.update_correction_table()
        if self.bad_weights.empty:
            QMessageBox.information(self, "Info", "No issues found with weights.")
        logger.debug(f"Check weights took {time.time() - start_time:.3f} seconds")

    def update_correction_table(self):
        """Update the correction table with bad weights."""
        start_time = time.time()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Exclude", "Solution Label", "Act Wgt"])

        if self.bad_weights is not None and not self.bad_weights.empty:
            self.correction_weight = {}
            excluded_samples = self.app.get_excluded_samples()
            for _, row in self.bad_weights.iterrows():
                solution_label = row['Solution Label']
                actual_weight = row['Act Wgt']

                # Checkbox for exclusion using PyQt6 native checkbox
                exclude_item = QStandardItem()
                exclude_item.setCheckable(True)
                exclude_item.setCheckState(Qt.CheckState.Checked if solution_label in excluded_samples else Qt.CheckState.Unchecked)
                exclude_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                label_item = QStandardItem(str(solution_label))
                label_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft)
                weight_item = QStandardItem(f"{actual_weight:.3f}")
                weight_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)

                model.appendRow([exclude_item, label_item, weight_item])
                self.correction_weight[solution_label] = self.new_weight_entry

            # Connect checkbox toggle
            model.itemChanged.connect(self.toggle_exclude)

        self.correction_table.setModel(model)
        self.correction_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(0, 80)  # Reduced width for checkbox
        self.correction_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.correction_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(2, 80)  # Reduced width for weight
        self.correction_table.setStyleSheet("QTableView { padding: 0; }")  # Remove extra padding

        logger.debug(f"Updating correction table took {time.time() - start_time:.3f} seconds")

    def toggle_exclude(self, item):
        """Toggle exclusion of a sample."""
        start_time = time.time()
        if item.column() == 0:
            model = self.correction_table.model()
            solution_label = model.data(model.index(item.row(), 1))
            new_state = item.checkState()
            if new_state == Qt.CheckState.Checked:
                self.app.add_excluded_sample(solution_label)
            else:
                self.app.remove_excluded_sample(solution_label)
            logger.debug(f"Toggle exclude for {solution_label} took {time.time() - start_time:.3f} seconds")

    def apply_weight_correction(self):
        """Apply weight correction to the selected solution label using a thread."""
        start_time = time.time()
        if not self.selected_solution_label:
            QMessageBox.warning(self, "Warning", "No solution label selected!")
            return

        try:
            new_weight = float(self.new_weight_entry.text())
            if new_weight <= 0:
                raise ValueError("Weight must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid weight: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading in apply_weight_correction took {time.time() - data_start:.3f} seconds")

        df = self.df_cache
        if df is None or df.empty:
            QMessageBox.warning(self, "Warning", "No data loaded!")
            return

        # Check if the selected row is excluded
        model = self.correction_table.model()
        for row in range(model.rowCount()):
            if model.data(model.index(row, 1)) == self.selected_solution_label:
                if model.item(row, 0).checkState() == Qt.CheckState.Checked:
                    QMessageBox.warning(self, "Warning", f"{self.selected_solution_label} is excluded and cannot be corrected!")
                    return
                break

        # Start thread
        self.thread = WeightCorrectionThread(df, [self.selected_solution_label], new_weight, apply_to_all=False)
        self.progress_dialog = QProgressDialog("Applying weight correction...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.canceled.connect(self.thread.terminate)
        self.thread.progress.connect(self.progress_dialog.setValue)
        self.thread.finished.connect(self.on_correction_finished)
        self.thread.error.connect(self.on_correction_error)
        self.thread.start()

        logger.debug(f"Starting apply_weight_correction took {time.time() - start_time:.3f} seconds")

    def apply_to_all(self):
        """Apply the new weight to all non-excluded samples in the bad weights table using a thread."""
        start_time = time.time()
        if self.bad_weights is None or self.bad_weights.empty:
            QMessageBox.warning(self, "Warning", "No bad weights to correct!")
            return

        try:
            new_weight = float(self.new_weight_entry.text())
            if new_weight <= 0:
                raise ValueError("Weight must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid weight: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading in apply_to_all took {time.time() - data_start:.3f} seconds")

        df = self.df_cache
        if df is None or df.empty:
            QMessageBox.warning(self, "Warning", "No data loaded!")
            return

        # Filter out excluded samples
        model = self.correction_table.model()
        non_excluded_labels = []
        for row in range(model.rowCount()):
            if model.item(row, 0).checkState() != Qt.CheckState.Checked:
                non_excluded_labels.append(model.data(model.index(row, 1)))

        if not non_excluded_labels:
            QMessageBox.warning(self, "Warning", "All samples are excluded!")
            return

        # Start thread
        self.thread = WeightCorrectionThread(df, non_excluded_labels, new_weight, apply_to_all=True)
        self.progress_dialog = QProgressDialog("Applying weight corrections to all samples...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.canceled.connect(self.thread.terminate)
        self.thread.progress.connect(self.progress_dialog.setValue)
        self.thread.finished.connect(self.on_correction_finished)
        self.thread.error.connect(self.on_correction_error)
        self.thread.start()

        logger.debug(f"Starting apply_to_all took {time.time() - start_time:.3f} seconds")

    def on_correction_finished(self, df_json, corrected_rows):
        """Handle thread completion."""
        self.df_cache = pd.read_json(df_json)
        self.app.set_data(self.df_cache)
        self.app.notify_data_changed()
        self.bad_weights = None
        self.check_weights()
        QMessageBox.information(self, "Success", f"Corrected weights and Corr Con values for {corrected_rows} rows")
        self.progress_dialog.close()

    def on_correction_error(self, error_msg):
        """Handle thread errors."""
        QMessageBox.warning(self, "Error", f"Failed to apply corrections: {error_msg}")
        self.progress_dialog.close()