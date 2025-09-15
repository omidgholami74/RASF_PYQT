from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont
import pandas as pd
import numpy as np
import time
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

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
        """Set up the UI with controls and a scrollable table."""
        start_time = time.time()
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
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Input frame
        input_frame = QFrame()
        input_layout = QHBoxLayout(input_frame)
        input_layout.setSpacing(10)

        input_layout.addWidget(QLabel("Expected Volume:"))
        self.volume_entry = QLineEdit()
        self.volume_entry.setText(str(self.volume_value))
        self.volume_entry.setFixedWidth(100)
        input_layout.addWidget(self.volume_entry)

        check_button = QPushButton("Check Volumes")
        check_button.clicked.connect(self.check_volumes)
        input_layout.addWidget(check_button)

        main_layout.addWidget(input_frame)

        # Main container
        main_container = QFrame()
        main_layout.addWidget(main_container, stretch=1)
        container_layout = QHBoxLayout(main_container)
        container_layout.setSpacing(10)

        # Correction frame
        correction_frame = QFrame()
        correction_layout = QVBoxLayout(correction_frame)
        correction_layout.setSpacing(10)

        # New volume input
        new_volume_frame = QFrame()
        new_volume_layout = QHBoxLayout(new_volume_frame)
        new_volume_layout.setSpacing(5)
        new_volume_layout.addWidget(QLabel("New Volume:"))
        self.new_volume_entry = QLineEdit()
        self.new_volume_entry.setText(str(self.new_volume))
        self.new_volume_entry.setFixedWidth(100)
        new_volume_layout.addWidget(self.new_volume_entry)
        correction_button = QPushButton("Apply Correction")
        correction_button.clicked.connect(self.apply_volume_correction)
        new_volume_layout.addWidget(correction_button)
        correction_layout.addWidget(new_volume_frame)

        # Bad volumes table
        self.correction_table = QTableView()
        self.correction_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.correction_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.correction_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.correction_table.verticalHeader().setVisible(False)
        self.correction_table.clicked.connect(self.select_row)
        correction_layout.addWidget(self.correction_table)

        container_layout.addWidget(correction_frame, stretch=1)

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

                # Checkbox for exclusion
                exclude_item = QStandardItem()
                exclude_item.setCheckable(True)
                exclude_item.setCheckState(Qt.CheckState.Checked if solution_label in excluded_volumes else Qt.CheckState.Unchecked)
                exclude_item.setText("☑" if solution_label in excluded_volumes else "☐")
                exclude_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                label_item = QStandardItem(str(solution_label))
                volume_item = QStandardItem(f"{actual_volume:.3f}")

                model.appendRow([exclude_item, label_item, volume_item])
                self.correction_volume[solution_label] = self.new_volume_entry

            # Connect checkbox toggle
            model.itemChanged.connect(self.toggle_exclude)

        self.correction_table.setModel(model)
        self.correction_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(0, 80)
        self.correction_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.correction_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.correction_table.setColumnWidth(2, 100)

        logger.debug(f"Updating correction table took {time.time() - start_time:.3f} seconds")

    def toggle_exclude(self, item):
        """Toggle exclusion of a volume."""
        start_time = time.time()
        if item.column() == 0:
            model = self.correction_table.model()
            solution_label = model.data(model.index(item.row(), 1))
            new_state = item.checkState()
            item.setText("☑" if new_state == Qt.CheckState.Checked else "☐")
            if new_state == Qt.CheckState.Checked:
                self.app.add_excluded_volume(solution_label)
            else:
                self.app.remove_excluded_volume(solution_label)
            logger.debug(f"Toggle exclude for {solution_label} took {time.time() - start_time:.3f} seconds")

    def apply_volume_correction(self):
        """Apply volume correction to the selected solution label."""
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
        mask = (df['Solution Label'] == self.selected_solution_label) & (df['Type'] == 'Samp')
        matching_rows = df[mask]
        if matching_rows.empty:
            QMessageBox.warning(self, "Warning", f"No matching rows found for {self.selected_solution_label}")
            return

        correction_start = time.time()
        for idx in matching_rows.index:
            current_volume = float(df.loc[idx, 'Act Vol'])
            current_corr_con = float(df.loc[idx, 'Corr Con'])
            corrected_corr_con = (self.new_volume / current_volume) * current_corr_con
            df.loc[idx, 'Corr Con'] = corrected_corr_con
            df.loc[idx, 'Act Vol'] = self.new_volume

        self.app.set_data(df)
        self.app.notify_data_changed()
        self.bad_volumes = None
        self.check_volumes()
        QMessageBox.information(self, "Success", f"Corrected {self.selected_solution_label} volume and Corr Con values for {len(matching_rows)} rows")
        logger.debug(f"Apply volume correction took {time.time() - correction_start:.3f} seconds")
        logger.debug(f"Total apply_volume_correction took {time.time() - start_time:.3f} seconds")