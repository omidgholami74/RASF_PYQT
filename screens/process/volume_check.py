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

class VolumeCorrectionThread(QThread):
    """Thread for applying volume corrections in the background."""
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, int)
    error = pyqtSignal(str)

    def __init__(self, df, solution_labels, new_volume, apply_to_all=False):
        super().__init__()
        self.df = df.copy()
        self.solution_labels = solution_labels
        self.new_volume = new_volume
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
                        current_volume = float(self.df.loc[idx, 'Act Vol'])
                        current_corr_con = float(self.df.loc[idx, 'Corr Con'])
                        corrected_corr_con = (self.new_volume / current_volume) * current_corr_con
                        self.df.loc[idx, 'Corr Con'] = corrected_corr_con
                        self.df.loc[idx, 'Act Vol'] = self.new_volume
                    corrected_rows += len(matching_rows)
                if self.apply_to_all:
                    self.progress.emit(int((i + 1) / total_rows * 100))
            self.finished.emit(self.df.to_json(), corrected_rows)
        except Exception as e:
            self.error.emit(str(e))

class VolumeCheckFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.df_cache = None
        self.bad_volumes = None
        self.correction_volume = {}
        self.selected_solution_label = None
        self.volume_value = 50.0
        self.new_volume = 50.0
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
        input_group = QGroupBox("Volume Check")
        input_layout = QHBoxLayout(input_group)
        input_layout.setSpacing(10)

        input_layout.addWidget(QLabel("Expected Volume:"))
        self.volume_entry = QLineEdit()
        self.volume_entry.setText(str(self.volume_value))
        self.volume_entry.setFixedWidth(120)
        self.volume_entry.setToolTip("Enter the expected volume (e.g., 50.0)")
        input_layout.addWidget(self.volume_entry)

        check_button = QPushButton("Check Volumes")
        check_button.setToolTip("Check volumes against the expected value")
        check_button.clicked.connect(self.check_volumes)
        input_layout.addWidget(check_button)
        input_layout.addStretch()

        main_layout.addWidget(input_group)

        # Main container
        main_container = QFrame()
        main_layout.addWidget(main_container, stretch=1)
        container_layout = QHBoxLayout(main_container)
        container_layout.setSpacing(15)

        # Correction group
        correction_group = QGroupBox("Volume Correction")
        correction_layout = QVBoxLayout(correction_group)
        correction_layout.setSpacing(10)

        # New volume input
        new_volume_frame = QFrame()
        new_volume_layout = QHBoxLayout(new_volume_frame)
        new_volume_layout.setSpacing(10)
        new_volume_layout.addWidget(QLabel("New Volume:"))
        self.new_volume_entry = QLineEdit()
        self.new_volume_entry.setText(str(self.new_volume))
        self.new_volume_entry.setFixedWidth(120)
        self.new_volume_entry.setToolTip("Enter the new volume to apply to the selected sample")
        new_volume_layout.addWidget(self.new_volume_entry)

        correction_button = QPushButton("Apply Correction")
        correction_button.setToolTip("Apply the new volume to the selected sample")
        correction_button.clicked.connect(self.apply_volume_correction)
        new_volume_layout.addWidget(correction_button)

        apply_all_button = QPushButton("Apply to All")
        apply_all_button.setToolTip("Apply the new volume to all non-excluded samples in the table")
        apply_all_button.clicked.connect(self.apply_to_all)
        new_volume_layout.addWidget(apply_all_button)

        new_volume_layout.addStretch()
        correction_layout.addWidget(new_volume_frame)

        # Bad volumes table
        self.correction_table = QTableView()
        self.correction_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.correction_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.correction_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.correction_table.verticalHeader().setVisible(False)
        self.correction_table.clicked.connect(self.select_row)
        self.correction_table.setToolTip("Select a row to correct its volume or mark it for exclusion")
        correction_layout.addWidget(self.correction_table)

        container_layout.addWidget(correction_group, stretch=1)

        logger.debug(f"UI setup took {time.time() - start_time:.3f} seconds")

    def select_row(self, index):
        """Handle row selection in the correction table."""
        if not index.isValid():
            return
        model = self.correction_table.model()
        self.selected_solution_label = model.data(model.index(index.row(), 1))
        if self.bad_volumes is not None:
            volume_row = self.bad_volumes[self.bad_volumes['Solution Label'] == self.selected_solution_label]
            if not volume_row.empty:
                try:
                    actual_volume = float(volume_row['Act Vol'].iloc[0])
                    self.new_volume_entry.setText(f"{actual_volume:.3f}")
                except (ValueError, TypeError) as e:
                    logger.error(f"Error setting new volume: {e}")
                    self.new_volume_entry.setText("50.0")

    def check_volumes(self):
        """Check volumes and display bad volumes in the table."""
        start_time = time.time()
        try:
            self.volume_value = float(self.volume_entry.text())
            if self.volume_value <= 0:
                raise ValueError("Expected volume must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid volume: {e}")
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

        # Filter bad volumes
        data_filter_start = time.time()
        sample_data = df[df['Type'] == 'Samp']
        self.bad_volumes = sample_data[
            (sample_data['Act Vol'] != self.volume_value)
        ][['Solution Label', 'Act Vol', 'Corr Con']].drop_duplicates(subset=['Solution Label'])
        logger.debug(f"Filtering bad volumes took {time.time() - data_filter_start:.3f} seconds")

        # Update table
        self.update_correction_table()
        if self.bad_volumes.empty:
            QMessageBox.information(self, "Info", "No issues found with volumes.")
        logger.debug(f"Check volumes took {time.time() - start_time:.3f} seconds")

    def update_correction_table(self):
        """Update the correction table with bad volumes."""
        start_time = time.time()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Exclude", "Solution Label", "Act Vol"])

        if self.bad_volumes is not None and not self.bad_volumes.empty:
            self.correction_volume = {}
            excluded_volumes = self.app.get_excluded_volumes()
            for _, row in self.bad_volumes.iterrows():
                solution_label = row['Solution Label']
                actual_volume = row['Act Vol']

                # Checkbox for exclusion using PyQt6 native checkbox
                exclude_item = QStandardItem()
                exclude_item.setCheckable(True)
                exclude_item.setCheckState(Qt.CheckState.Checked if solution_label in excluded_volumes else Qt.CheckState.Unchecked)
                exclude_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                label_item = QStandardItem(str(solution_label))
                label_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft)
                volume_item = QStandardItem(f"{actual_volume:.3f}")
                volume_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)

                model.appendRow([exclude_item, label_item, volume_item])
                self.correction_volume[solution_label] = self.new_volume_entry

            # Connect checkbox toggle
            model.itemChanged.connect(self.toggle_exclude)

        self.correction_table.setModel(model)
        self.correction_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(0, 80)  # Reduced width for checkbox
        self.correction_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.correction_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(2, 80)  # Reduced width for volume
        self.correction_table.setStyleSheet("QTableView { padding: 0; }")  # Remove extra padding

        logger.debug(f"Updating correction table took {time.time() - start_time:.3f} seconds")

    def toggle_exclude(self, item):
        """Toggle exclusion of a volume."""
        start_time = time.time()
        if item.column() == 0:
            model = self.correction_table.model()
            solution_label = model.data(model.index(item.row(), 1))
            new_state = item.checkState()
            if new_state == Qt.CheckState.Checked:
                self.app.add_excluded_volume(solution_label)
            else:
                self.app.remove_excluded_volume(solution_label)
            logger.debug(f"Toggle exclude for {solution_label} took {time.time() - start_time:.3f} seconds")

    def apply_volume_correction(self):
        """Apply volume correction to the selected solution label using a thread."""
        start_time = time.time()
        if not self.selected_solution_label:
            QMessageBox.warning(self, "Warning", "No solution label selected!")
            return

        try:
            self.new_volume = float(self.new_volume_entry.text())
            if self.new_volume <= 0:
                raise ValueError("Volume must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid volume: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading in apply_volume_correction took {time.time() - data_start:.3f} seconds")

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
        self.thread = VolumeCorrectionThread(df, [self.selected_solution_label], self.new_volume, apply_to_all=False)
        self.progress_dialog = QProgressDialog("Applying volume correction...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.canceled.connect(self.thread.terminate)
        self.thread.progress.connect(self.progress_dialog.setValue)
        self.thread.finished.connect(self.on_correction_finished)
        self.thread.error.connect(self.on_correction_error)
        self.thread.start()

        logger.debug(f"Starting apply_volume_correction took {time.time() - start_time:.3f} seconds")

    def apply_to_all(self):
        """Apply the new volume to all non-excluded samples in the bad volumes table using a thread."""
        start_time = time.time()
        if self.bad_volumes is None or self.bad_volumes.empty:
            QMessageBox.warning(self, "Warning", "No bad volumes to correct!")
            return

        try:
            self.new_volume = float(self.new_volume_entry.text())
            if self.new_volume <= 0:
                raise ValueError("Volume must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid volume: {e}")
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
        self.thread = VolumeCorrectionThread(df, non_excluded_labels, self.new_volume, apply_to_all=True)
        self.progress_dialog = QProgressDialog("Applying volume corrections to all samples...", "Cancel", 0, 100, self)
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
        self.bad_volumes = None
        self.check_volumes()
        QMessageBox.information(self, "Success", f"Corrected volumes and Corr Con values for {corrected_rows} rows")
        self.progress_dialog.close()

    def on_correction_error(self, error_msg):
        """Handle thread errors."""
        QMessageBox.warning(self, "Error", f"Failed to apply corrections: {error_msg}")
        self.progress_dialog.close()