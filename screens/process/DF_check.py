from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit, QPushButton, QTableView, QHeaderView, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import pandas as pd
import re
import time
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class DFCheckFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.df_cache = None
        self.bad_dfs = None
        self.correction_df = {}
        self.selected_solution_label = None
        self.df_value = 1.0
        self.new_df = 1.0
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

        input_layout.addWidget(QLabel("Expected DF Value:"))
        self.df_entry = QLineEdit()
        self.df_entry.setText(str(self.df_value))
        self.df_entry.setFixedWidth(100)
        input_layout.addWidget(self.df_entry)

        check_button = QPushButton("Check DF Values")
        check_button.clicked.connect(self.check_df_values)
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

        # New DF input
        new_df_frame = QFrame()
        new_df_layout = QHBoxLayout(new_df_frame)
        new_df_layout.setSpacing(5)
        new_df_layout.addWidget(QLabel("New DF:"))
        self.new_df_entry = QLineEdit()
        self.new_df_entry.setText(str(self.new_df))
        self.new_df_entry.setFixedWidth(100)
        new_df_layout.addWidget(self.new_df_entry)
        correction_button = QPushButton("Apply Correction")
        correction_button.clicked.connect(self.apply_df_correction)
        new_df_layout.addWidget(correction_button)
        correction_layout.addWidget(new_df_frame)

        # Bad DFs table
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
        if self.bad_dfs is not None:
            df_row = self.bad_dfs[self.bad_dfs['Solution Label'] == self.selected_solution_label]
            if not df_row.empty:
                try:
                    actual_df = float(df_row['DF'].iloc[0])
                    self.new_df_entry.setText(f"{actual_df:.3f}")
                except (ValueError, TypeError) as e:
                    logger.error(f"Error setting new DF: {e}")
                    self.new_df_entry.setText("1.0")

    def check_df_values(self):
        """Check samples where DF doesn't match the number after 'D' in Solution Label or expected input."""
        start_time = time.time()
        try:
            self.df_value = float(self.df_entry.text())
            if self.df_value <= 0:
                raise ValueError("Expected DF must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid DF: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading took {time.time() - data_start:.3f} seconds")

        df = self.df_cache
        if df is None or df.empty:
            QMessageBox.warning(self, "Warning", "No data loaded!")
            return

        sample_data = df[df['Type'] == 'Samp'].copy()
        if sample_data.empty:
            QMessageBox.warning(self, "Warning", "No sample data found!")
            return

        # Extract expected DF from Solution Label or use input
        def get_expected_df(label):
            match = re.search(r'D(\d+)', label)
            return int(match.group(1)) if match else self.df_value

        data_filter_start = time.time()
        sample_data['Expected DF'] = sample_data['Solution Label'].apply(get_expected_df)
        sample_data['DF'] = pd.to_numeric(sample_data['DF'], errors='coerce')
        self.bad_dfs = sample_data[
            (sample_data['DF'] != sample_data['Expected DF'])
        ][['Solution Label', 'DF']].drop_duplicates(subset=['Solution Label'])
        logger.debug(f"Filtering bad DFs took {time.time() - data_filter_start:.3f} seconds")

        # Update table
        self.update_correction_table()
        if self.bad_dfs.empty:
            QMessageBox.information(self, "Info", "No issues found with DF values.")
        logger.debug(f"Check DF values took {time.time() - start_time:.3f} seconds")

    def update_correction_table(self):
        """Update the correction table with bad DFs."""
        start_time = time.time()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Exclude", "Solution Label", "DF"])

        if self.bad_dfs is not None and not self.bad_dfs.empty:
            self.correction_df = {}
            excluded_dfs = self.app.get_excluded_dfs()
            for _, row in self.bad_dfs.iterrows():
                solution_label = row['Solution Label']
                actual_df = row['DF']

                # Checkbox for exclusion
                exclude_item = QStandardItem()
                exclude_item.setCheckable(True)
                exclude_item.setCheckState(Qt.CheckState.Checked if solution_label in excluded_dfs else Qt.CheckState.Unchecked)
                exclude_item.setText("☑" if solution_label in excluded_dfs else "☐")
                exclude_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                label_item = QStandardItem(str(solution_label))
                df_item = QStandardItem(f"{actual_df:.3f}")

                model.appendRow([exclude_item, label_item, df_item])
                self.correction_df[solution_label] = self.new_df_entry

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
        """Toggle exclusion of a DF."""
        start_time = time.time()
        if item.column() == 0:
            model = self.correction_table.model()
            solution_label = model.data(model.index(item.row(), 1))
            new_state = item.checkState()
            item.setText("☑" if new_state == Qt.CheckState.Checked else "☐")
            if new_state == Qt.CheckState.Checked:
                self.app.add_excluded_df(solution_label)
            else:
                self.app.remove_excluded_df(solution_label)
            self.app.notify_data_changed()
            logger.debug(f"Toggle exclude for {solution_label} took {time.time() - start_time:.3f} seconds")

    def apply_df_correction(self):
        """Apply DF correction to the selected solution label."""
        start_time = time.time()
        if not self.selected_solution_label:
            QMessageBox.warning(self, "Warning", "No solution label selected!")
            return

        try:
            self.new_df = float(self.new_df_entry.text())
            if self.new_df <= 0:
                raise ValueError("DF must be positive")
        except ValueError as e:
            QMessageBox.warning(self, "Warning", f"Invalid DF: {e}")
            return

        if self.df_cache is None:
            data_start = time.time()
            self.df_cache = self.app.get_data()
            logger.debug(f"Data loading in apply_df_correction took {time.time() - data_start:.3f} seconds")

        df = self.df_cache
        mask = (df['Solution Label'] == self.selected_solution_label) & (df['Type'] == 'Samp')
        matching_rows = df[mask]
        if matching_rows.empty:
            QMessageBox.warning(self, "Warning", f"No matching rows found for {self.selected_solution_label}")
            return

        correction_start = time.time()
        df.loc[mask, 'DF'] = self.new_df

        self.app.set_data(df)
        self.app.notify_data_changed()
        self.bad_dfs = None
        self.check_df_values()
        QMessageBox.information(self, "Success", f"Corrected {self.selected_solution_label} DF value for {len(matching_rows)} rows")
        logger.debug(f"Apply DF correction took {time.time() - correction_start:.3f} seconds")
        logger.debug(f"Total apply_df_correction took {time.time() - start_time:.3f} seconds")