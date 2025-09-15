from PyQt6.QtWidgets import (QWidget, QFrame, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QTableView, QHeaderView, QMessageBox)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QStandardItemModel, QStandardItem
import pandas as pd
import numpy as np
import json
import os
import time
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class CheckRMFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.rm_df = None
        self.positions_df = None
        self.original_df = None
        self.outliers = {}
        self.current_label = None
        self.corrected_df = None
        self.selected_element = None
        self.selected_row_id_pair = None
        self.corrections_applied = False
        self.ratios = {}
        self.non_outlier_ratios = {}
        self.selected_outliers = {}
        self.ignored_outliers = {}
        self.user_corrections = {}
        self.corrections_file = "user_corrections.json"
        self.load_user_corrections()
        self.setup_ui()

    def configure_style(self):
        """Apply consistent styling to the frame."""
        self.setStyleSheet("""
            QWidget {
                background-color: #FAFAFA;
                border: none;
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #B0BEC5;
                padding: 5px;
                font: 12px 'Segoe UI';
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
            QHeaderView::section {
                background-color: #ECEFF1;
                font: bold 12px 'Segoe UI';
                border: 1px solid #ccc;
                padding: 5px;
            }
            QTableView {
                background-color: white;
                gridline-color: #ccc;
                font: 11px 'Segoe UI';
            }
            QTableView::item {
                padding: 2px;
                border: none;
            }
            QTableView::item:selected {
                background-color: #BBDEFB;
                color: black;
            }
        """)

    def load_user_corrections(self):
        """Load user corrections from JSON file."""
        if os.path.exists(self.corrections_file):
            try:
                with open(self.corrections_file, 'r') as f:
                    self.user_corrections = json.load(f)
                logger.info("User corrections loaded successfully")
            except Exception as e:
                logger.error(f"Error loading user corrections: {e}")

    def save_user_corrections(self):
        """Save user corrections to JSON file."""
        try:
            with open(self.corrections_file, 'w') as f:
                json.dump(self.user_corrections, f, indent=4)
            logger.info("User corrections saved successfully")
        except Exception as e:
            logger.error(f"Error saving user corrections: {e}")

    def setup_ui(self):
        """Setup UI with controls, tables, and plot area."""
        self.configure_style()
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Control frame
        control_frame = QFrame()
        control_layout = QHBoxLayout(control_frame)
        control_layout.setSpacing(15)

        keyword_label = QLabel("Keyword:")
        keyword_label.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        control_layout.addWidget(keyword_label)
        self.keyword_entry = QLineEdit()
        self.keyword_entry.setText("RM")
        self.keyword_entry.setFixedWidth(100)
        control_layout.addWidget(self.keyword_entry)

        threshold_label = QLabel("Threshold (σ):")
        threshold_label.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        control_layout.addWidget(threshold_label)
        self.threshold_entry = QLineEdit()
        self.threshold_entry.setText("2.0")
        self.threshold_entry.setFixedWidth(100)
        control_layout.addWidget(self.threshold_entry)

        self.run_button = QPushButton("Check RM Changes")
        self.run_button.clicked.connect(self.check_rm_changes)
        control_layout.addWidget(self.run_button)

        self.apply_correction_button = QPushButton("Apply Ratio Correction")
        self.apply_correction_button.clicked.connect(self.apply_ratio_correction)
        self.apply_correction_button.setEnabled(False)
        control_layout.addWidget(self.apply_correction_button)

        self.apply_all_button = QPushButton("Apply All Selected")
        self.apply_all_button.clicked.connect(self.apply_all_corrections)
        self.apply_all_button.setEnabled(False)
        control_layout.addWidget(self.apply_all_button)

        self.skip_outlier_button = QPushButton("Skip Selected Outlier")
        self.skip_outlier_button.clicked.connect(self.skip_selected_outlier)
        self.skip_outlier_button.setEnabled(False)
        control_layout.addWidget(self.skip_outlier_button)

        self.mark_outlier_button = QPushButton("Mark as Outlier")
        self.mark_outlier_button.clicked.connect(self.mark_as_outlier)
        self.mark_outlier_button.setEnabled(False)
        control_layout.addWidget(self.mark_outlier_button)

        self.mark_non_outlier_button = QPushButton("Mark as Non-Outlier")
        self.mark_non_outlier_button.clicked.connect(self.mark_as_non_outlier)
        self.mark_non_outlier_button.setEnabled(False)
        control_layout.addWidget(self.mark_non_outlier_button)

        main_layout.addWidget(control_frame)

        # Content frame
        content_frame = QFrame()
        content_layout = QHBoxLayout(content_frame)
        content_layout.setSpacing(15)

        # Left frame for tables
        left_frame = QFrame()
        left_layout = QVBoxLayout(left_frame)
        left_layout.setSpacing(15)

        # Outliers table
        outliers_label = QLabel("Outliers Table")
        outliers_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        left_layout.addWidget(outliers_label)

        self.outliers_table = QTableView()
        self.outliers_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.outliers_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.outliers_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.outliers_table.verticalHeader().setVisible(False)
        self.outliers_table.clicked.connect(self.on_outlier_select)
        left_layout.addWidget(self.outliers_table)

        # Ratios table
        ratios_label = QLabel("Outlier Points and Ratios")
        ratios_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        left_layout.addWidget(ratios_label)

        self.ratios_table = QTableView()
        self.ratios_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.ratios_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.ratios_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.ratios_table.verticalHeader().setVisible(False)
        self.ratios_table.clicked.connect(self.on_ratio_select)
        left_layout.addWidget(self.ratios_table)

        # Non-outlier table
        non_outlier_label = QLabel("Non-Outlier Points and Ratios")
        non_outlier_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        left_layout.addWidget(non_outlier_label)

        self.non_outlier_table = QTableView()
        self.non_outlier_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.non_outlier_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.non_outlier_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.non_outlier_table.verticalHeader().setVisible(False)
        self.non_outlier_table.clicked.connect(self.on_non_outlier_select)
        left_layout.addWidget(self.non_outlier_table)

        content_layout.addWidget(left_frame, stretch=1)

        # Right frame for plots and preview
        right_frame = QFrame()
        right_layout = QVBoxLayout(right_frame)
        right_layout.setSpacing(15)

        plot_label = QLabel("Trend Plots")
        plot_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        right_layout.addWidget(plot_label)

        self.plot_frame = QFrame()
        self.plot_frame.setStyleSheet("background-color: white; border: 1px solid #E0E0E0; border-radius: 10px;")
        self.plot_frame.setMinimumHeight(300)
        self.plot_layout = QVBoxLayout(self.plot_frame)
        right_layout.addWidget(self.plot_frame)

        corrected_label = QLabel("Between Elements Preview")
        corrected_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        right_layout.addWidget(corrected_label)

        self.corrected_table = QTableView()
        self.corrected_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.corrected_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.corrected_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.corrected_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.corrected_table)

        content_layout.addWidget(right_frame, stretch=1)
        main_layout.addWidget(content_frame)

    def check_rm_changes(self):
        """Check for RM changes and detect outliers."""
        start_time = time.time()
        try:
            threshold_sigma = float(self.threshold_entry.text())
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter a valid number for threshold (σ).")
            return

        keyword = self.keyword_entry.text().strip()
        if not keyword:
            QMessageBox.critical(self, "Error", "Please enter a valid keyword.")
            return

        df = self.app.get_data()
        if df is None or df.empty:
            QMessageBox.critical(self, "Error", "No data loaded.")
            return

        # Verify required columns in input DataFrame
        required_columns = ['Solution Label', 'Element', 'Type', 'Corr Con', 'Act Wgt', 'Act Vol', 'Coeff 1', 'Coeff 2']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            QMessageBox.critical(self, "Error", f"Missing required columns: {missing_columns}")
            return

        # Create original_df and ensure original_index
        self.original_df = df.copy(deep=True)
        self.original_df.reset_index(inplace=True)
        self.original_df.rename(columns={'index': 'original_index'}, inplace=True)

        df_filtered = df[df['Type'] == 'Samp'].copy(deep=True)
        if df_filtered.empty:
            QMessageBox.critical(self, "Error", f"No data with Type='Samp' found. Number of rows: {len(df)}")
            return

        df_filtered.reset_index(inplace=True)
        df_filtered.rename(columns={'index': 'original_index'}, inplace=True)
        df_filtered['Solution Label'] = df_filtered['Solution Label'].str.replace(
            rf'^{keyword}\s*[-]?\s*(\d*)$', rf'{keyword}\1', regex=True
        )
        df_filtered['row_id'] = df_filtered.groupby(['Solution Label', 'Element']).cumcount()

        # Merge row_id into original_df
        self.original_df = self.original_df.merge(
            df_filtered[['original_index', 'Solution Label', 'Element', 'row_id']],
            on=['original_index', 'Solution Label', 'Element'],
            how='left'
        )
        self.original_df['row_id'] = self.original_df['row_id'].fillna(-1).astype(int)

        # Convert columns to numeric
        numeric_columns = ['Corr Con', 'Act Wgt', 'Act Vol', 'Coeff 1', 'Coeff 2']
        for col in numeric_columns:
            df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce')
            self.original_df[col] = pd.to_numeric(self.original_df[col], errors='coerce')

        self.corrected_df = df_filtered.copy(deep=True)
        self.positions_df = df_filtered.groupby(['Solution Label', 'row_id'])['original_index'].agg(['min', 'max']).reset_index()
        pivot_df = df_filtered.pivot(
            index=['Solution Label', 'row_id'],
            columns='Element',
            values='Corr Con'
        ).reset_index()

        pivot_df['Solution Label'] = pivot_df['Solution Label'].fillna('')
        self.rm_df = pivot_df[pivot_df['Solution Label'].str.match(rf'^{keyword}\d*$')].copy(deep=True)

        if self.rm_df.empty:
            unique_labels = list(df_filtered['Solution Label'].unique())
            QMessageBox.critical(self, "Error", f"No data with {keyword} label found. Solution Labels: {unique_labels[:10]}{'...' if len(unique_labels) > 10 else ''}")
            return

        columns_to_check = [col for col in self.rm_df.columns if col not in ['Solution Label', 'row_id']]
        if not columns_to_check:
            QMessageBox.critical(self, "Error", "No Element columns found.")
            return

        for col in columns_to_check:
            self.rm_df[col] = pd.to_numeric(self.rm_df[col], errors='coerce')

        self.corrections_applied = False
        self.outliers = {}
        self.ratios = {}
        self.non_outlier_ratios = {}
        self.ignored_outliers = {}
        solution_labels = sorted(self.rm_df['Solution Label'].unique(), key=lambda x: int(x.replace(keyword, '')) if x.replace(keyword, '').isdigit() else 0)
        outlier_results = []

        for label in solution_labels:
            label_df = self.rm_df[self.rm_df['Solution Label'] == label].sort_values('row_id')
            if label_df.empty or len(label_df) < 2:
                continue
            self.outliers[label] = {}
            self.ratios[label] = {}
            self.non_outlier_ratios[label] = {}
            self.ignored_outliers[f"{label}"] = set()

            row_ids = label_df['row_id'].values
            for col in columns_to_check:
                values = pd.to_numeric(label_df[col], errors='coerce').values
                valid_mask = ~np.isnan(values)
                if np.sum(valid_mask) < 2:
                    continue

                valid_values = values[valid_mask]
                valid_row_ids = row_ids[valid_mask]

                user_key = f"{label}:{col}"
                user_outliers = set(self.user_corrections.get(user_key, {}).get('outliers', []))
                user_non_outliers = set(self.user_corrections.get(user_key, {}).get('non_outliers', []))
                ignored_outliers = set(self.ignored_outliers.get(label, set()))

                coefficients = np.polyfit(valid_row_ids, valid_values, 1)
                trendline = np.polyval(coefficients, valid_row_ids)
                residuals = valid_values - trendline
                residual_std = np.std(residuals) if len(residuals) > 0 else 1.0
                outlier_mask = np.abs(residuals) > threshold_sigma * residual_std
                outlier_mask |= np.isin(valid_row_ids, list(user_outliers))
                outlier_mask &= ~np.isin(valid_row_ids, list(user_non_outliers))
                outlier_mask |= np.isin(valid_row_ids, list(ignored_outliers))

                outlier_count = int(np.sum(outlier_mask))
                if outlier_count > 0:
                    self.outliers[label][col] = outlier_count
                    outlier_results.append({
                        'Solution Label': label,
                        'Element': col,
                        'Outliers Count': outlier_count
                    })

                outlier_row_ids = valid_row_ids[outlier_mask]
                for oid in outlier_row_ids:
                    key = f"{label}:{col}:{oid}"
                    self.ratios[key] = np.nan

                non_outlier_values = valid_values[~outlier_mask]
                non_outlier_row_ids = valid_row_ids[~outlier_mask]
                if len(non_outlier_values) >= 2:
                    for i in range(1, len(non_outlier_values)):
                        old_val = non_outlier_values[i-1]
                        new_val = non_outlier_values[i]
                        if old_val == 0 or np.isnan(old_val) or np.isnan(new_val):
                            continue
                        ratio = old_val / new_val
                        key = f"{label}:{col}:{non_outlier_row_ids[i-1]}->{non_outlier_row_ids[i]}"
                        self.non_outlier_ratios[key] = ratio

        if outlier_results:
            results_df = pd.DataFrame(outlier_results)
            self.display_outliers(results_df)
        else:
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(["Select", "Solution Label", "Element", "Outliers Count"])
            self.outliers_table.setModel(model)
            self.selected_outliers.clear()

        if solution_labels and columns_to_check:
            self.current_label = solution_labels[0]
            self.selected_element = columns_to_check[0]
            self.apply_corrections_for_label(self.current_label, self.selected_element)
            self.display_ratios(self.current_label, self.selected_element)
            self.display_non_outlier_ratios(self.current_label, self.selected_element)
            self.plot_trend(self.selected_element)
            between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
            between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
            self.display_between_df(between_df)
            self.apply_correction_button.setEnabled(True)
            self.apply_all_button.setEnabled(True)
            self.skip_outlier_button.setEnabled(True)
            self.mark_outlier_button.setEnabled(True)
            self.mark_non_outlier_button.setEnabled(True)

        std_data = self.original_df[self.original_df['Type'] == 'Std'].copy(deep=True)
        updated_df = pd.concat([self.corrected_df, std_data], ignore_index=True)
        self.app.set_data(updated_df, for_results=True)

        logger.debug(f"Check RM changes took {time.time() - start_time:.3f} seconds")

    def display_outliers(self, df):
        """Display outliers in the outliers_table using QStandardItemModel."""
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Select", "Solution Label", "Element", "Outliers Count"])
        self.selected_outliers.clear()

        for index, row in df.iterrows():
            select_item = QStandardItem()
            select_item.setCheckable(True)
            select_item.setCheckState(Qt.CheckState.Unchecked)
            select_item.setText("☐")
            select_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            label_item = QStandardItem(str(row['Solution Label']))
            element_item = QStandardItem(str(row['Element']))
            count_item = QStandardItem(str(row['Outliers Count']))

            model.appendRow([select_item, label_item, element_item, count_item])
            self.selected_outliers[index] = select_item

        self.outliers_table.setModel(model)
        self.outliers_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.outliers_table.setColumnWidth(0, 60)
        self.outliers_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.outliers_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.outliers_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.outliers_table.setColumnWidth(3, 120)
        self.outliers_table.clicked.connect(self.toggle_checkbox)

    def toggle_checkbox(self, index):
        """Toggle checkbox state in the outliers table."""
        if index.column() == 0:
            model = self.outliers_table.model()
            item = model.item(index.row(), 0)
            if item.isCheckable():
                new_state = Qt.CheckState.Checked if item.checkState() == Qt.CheckState.Unchecked else Qt.CheckState.Unchecked
                item.setCheckState(new_state)
                item.setText("☑" if new_state == Qt.CheckState.Checked else "☐")

    def apply_corrections_for_label(self, label, element):
        """Apply corrections for a specific label and element."""
        label_df = self.rm_df[self.rm_df['Solution Label'] == label].sort_values('row_id')
        if label_df.empty:
            return

        row_ids = label_df['row_id'].values
        values = pd.to_numeric(label_df[element], errors='coerce').values
        valid_mask = ~np.isnan(values)
        if np.sum(valid_mask) < 2:
            return

        valid_values = values[valid_mask]
        valid_row_ids = row_ids[valid_mask]

        user_key = f"{label}:{element}"
        user_outliers = set(self.user_corrections.get(user_key, {}).get('outliers', []))
        user_non_outliers = set(self.user_corrections.get(user_key, {}).get('non_outliers', []))
        ignored_outliers = set(self.ignored_outliers.get(label, set()))

        coefficients = np.polyfit(valid_row_ids, valid_values, 1)
        trendline = np.polyval(coefficients, valid_row_ids)
        residuals = valid_values - trendline
        residual_std = np.std(residuals) if len(residuals) > 0 else 1.0
        outlier_mask = np.abs(residuals) > float(self.threshold_entry.text()) * residual_std
        outlier_mask |= np.isin(valid_row_ids, list(user_outliers))
        outlier_mask &= ~np.isin(valid_row_ids, list(user_non_outliers))
        outlier_mask |= np.isin(valid_row_ids, list(ignored_outliers))

        non_outlier_values = valid_values[~outlier_mask]
        non_outlier_row_ids = valid_row_ids[~outlier_mask]
        if len(non_outlier_values) > 0:
            mean_value = np.mean(non_outlier_values)
            condition = (self.corrected_df['Solution Label'] == label) & (self.corrected_df['Element'] == element)
            valid_rows = self.corrected_df[condition]['row_id'].isin(non_outlier_row_ids)
            self.corrected_df.loc[condition & valid_rows, 'Corr Con'] = mean_value

            std_data = self.original_df[self.original_df['Type'] == 'Std'].copy(deep=True)
            updated_df = pd.concat([self.corrected_df, std_data], ignore_index=True)
            self.app.set_data(updated_df, for_results=True)

    def get_non_outlier_condition(self, label, element, old_id, new_id):
        """Get condition for non-outlier rows between old_id and new_id."""
        max_old = self.positions_df[(self.positions_df['Solution Label'] == label) & (self.positions_df['row_id'] == old_id)]['max'].values
        min_new = self.positions_df[(self.positions_df['Solution Label'] == label) & (self.positions_df['row_id'] == new_id)]['min'].values
        if len(max_old) == 0 or len(min_new) == 0:
            return pd.Series(False, index=self.corrected_df.index)
        return (self.corrected_df['original_index'] > max_old[0]) & (self.corrected_df['original_index'] < min_new[0])

    def get_first_non_outlier_condition(self, label, element):
        """Get the first non-outlier condition for a label and element."""
        for key in self.non_outlier_ratios:
            parts = key.split(':')
            if len(parts) != 3:
                continue
            if parts[0] == label and parts[1] == element:
                old_id, new_id = map(int, parts[2].split('->'))
                self.selected_row_id_pair = f"{old_id}->{new_id}"
                return self.get_non_outlier_condition(label, element, old_id, new_id)
        return pd.Series(False, index=self.corrected_df.index)

    def display_ratios(self, label=None, element=None):
        """Display outlier points and ratios in the ratios_table."""
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Solution Label", "Element", "Outlier Point", "Ratio"])

        for key, ratio in self.ratios.items():
            parts = key.split(':')
            if len(parts) != 3:
                continue
            key_label, key_element, point = parts
            if label and element and (key_label != label or key_element != element):
                continue
            point_display = f"Outlier Point {point}"
            model.appendRow([
                QStandardItem(key_label),
                QStandardItem(key_element),
                QStandardItem(point_display),
                QStandardItem("N/A")
            ])

        self.ratios_table.setModel(model)
        self.ratios_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def display_non_outlier_ratios(self, label=None, element=None):
        """Display non-outlier points and ratios in the non_outlier_table."""
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Solution Label", "Element", "Non-Outlier Point", "Ratio"])

        for key, ratio in self.non_outlier_ratios.items():
            parts = key.split(':')
            if len(parts) != 3:
                continue
            key_label, key_element, point = parts
            if label and element and (key_label != label or key_element != element):
                continue
            point_display = f"Non-Outlier Point {point}"
            model.appendRow([
                QStandardItem(key_label),
                QStandardItem(key_element),
                QStandardItem(point_display),
                QStandardItem(f"{ratio:.3f}")
            ])

        self.non_outlier_table.setModel(model)
        self.non_outlier_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def display_between_df(self, df):
        """Display corrected DataFrame in the corrected_table."""
        model = QStandardItemModel()
        columns = ["Solution Label", "Element", "Previous Corr Con", "Corr Con", "Soln Conc", "Intense", 
                   "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)", "row_id"]
        model.setHorizontalHeaderLabels(columns)

        if df is None or df.empty:
            logger.debug("No data to display in Between Elements Preview table")
            self.corrected_table.setModel(model)
            return

        for _, row in df.iterrows():
            original_row = self.original_df[
                (self.original_df['Solution Label'] == row['Solution Label']) &
                (self.original_df['Element'] == row['Element']) &
                (self.original_df['original_index'] == row['original_index'])
            ]
            if original_row.empty:
                prev_corr_con = np.nan
                prev_intense = np.nan
                soln_conc = np.nan
                intense = np.nan
                soln_conc_no_rm = np.nan
                intense_no_rm = np.nan
            else:
                prev_corr_con = original_row['Corr Con'].iloc[0]
                try:
                    corr_con = float(row['Corr Con']) if pd.notna(row['Corr Con']) else np.nan
                    act_wgt = float(row['Act Wgt']) if pd.notna(row['Act Wgt']) else np.nan
                    act_vol = float(row['Act Vol']) if pd.notna(row['Act Vol']) and row['Act Vol'] != 0 else np.nan
                    coeff_1 = float(row['Coeff 1']) if pd.notna(row['Coeff 1']) else np.nan
                    coeff_2 = float(row['Coeff 2']) if pd.notna(row['Coeff 2']) else np.nan
                    prev_corr_con = float(prev_corr_con) if pd.notna(prev_corr_con) else np.nan

                    soln_conc = corr_con * act_wgt / act_vol if pd.notna(corr_con) and pd.notna(act_wgt) and pd.notna(act_vol) else np.nan
                    intense = soln_conc * coeff_2 + coeff_1 if pd.notna(soln_conc) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan
                    soln_conc_no_rm = prev_corr_con * act_wgt / act_vol if pd.notna(prev_corr_con) and pd.notna(act_wgt) and pd.notna(act_vol) else np.nan
                    intense_no_rm = soln_conc_no_rm * coeff_2 + coeff_1 if pd.notna(soln_conc_no_rm) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan
                    prev_intense = soln_conc_no_rm * coeff_2 + coeff_1 if pd.notna(soln_conc_no_rm) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan
                except (TypeError, ValueError):
                    soln_conc = np.nan
                    intense = np.nan
                    soln_conc_no_rm = np.nan
                    intense_no_rm = np.nan
                    prev_intense = np.nan

            model.appendRow([
                QStandardItem(str(row['Solution Label'])),
                QStandardItem(str(row['Element'])),
                QStandardItem(f"{prev_corr_con:.3f}" if pd.notna(prev_corr_con) else "N/A"),
                QStandardItem(f"{row['Corr Con']:.3f}" if pd.notna(row['Corr Con']) else "N/A"),
                QStandardItem(f"{soln_conc:.3f}" if pd.notna(soln_conc) else "N/A"),
                QStandardItem(f"{intense:.3f}" if pd.notna(intense) else "N/A"),
                QStandardItem(f"{prev_intense:.3f}" if pd.notna(prev_intense) else "N/A"),
                QStandardItem(f"{soln_conc_no_rm:.3f}" if pd.notna(soln_conc_no_rm) else "N/A"),
                QStandardItem(f"{intense_no_rm:.3f}" if pd.notna(intense_no_rm) else "N/A"),
                QStandardItem(str(row['row_id']))
            ])

        self.corrected_table.setModel(model)
        self.corrected_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def on_outlier_select(self, index):
        """Handle selection in the outliers table."""
        if not index.isValid():
            return
        row = index.row()
        model = self.outliers_table.model()
        self.current_label = model.data(model.index(row, 1))
        self.selected_element = model.data(model.index(row, 2))
        self.apply_corrections_for_label(self.current_label, self.selected_element)
        self.plot_trend(self.selected_element)
        self.display_ratios(self.current_label, self.selected_element)
        self.display_non_outlier_ratios(self.current_label, self.selected_element)
        between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)

    def on_ratio_select(self, index):
        """Handle selection in the ratios table."""
        if not index.isValid():
            return
        row = index.row()
        model = self.ratios_table.model()
        self.current_label = model.data(model.index(row, 0))
        self.selected_element = model.data(model.index(row, 1))
        point_display = model.data(model.index(row, 2))
        if point_display.startswith("Outlier Point"):
            self.selected_row_id_pair = None
        self.plot_trend(self.selected_element)
        between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)

    def on_non_outlier_select(self, index):
        """Handle selection in the non-outlier table."""
        if not index.isValid():
            return
        row = index.row()
        model = self.non_outlier_table.model()
        self.current_label = model.data(model.index(row, 0))
        self.selected_element = model.data(model.index(row, 1))
        point_display = model.data(model.index(row, 2))
        if point_display.startswith("Non-Outlier Point"):
            self.selected_row_id_pair = point_display.replace("Non-Outlier Point ", "")
        else:
            self.selected_row_id_pair = None
        between_condition = self.get_non_outlier_condition(self.current_label, self.selected_element, 
                                                          *map(int, self.selected_row_id_pair.split('->')) 
                                                          if self.selected_row_id_pair else (0, 0))
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)

    def apply_ratio_correction(self):
        """Apply ratio correction to selected non-outlier points."""
        if not self.selected_row_id_pair or not self.selected_element or not self.current_label:
            QMessageBox.critical(self, "Error", "Select a non-outlier point pair and element first.")
            return

        key = f"{self.current_label}:{self.selected_element}:{self.selected_row_id_pair}"
        if key not in self.non_outlier_ratios:
            QMessageBox.critical(self, "Error", "No ratio found for selection.")
            return

        ratio = self.non_outlier_ratios[key]
        old_id, new_id = map(int, self.selected_row_id_pair.split('->'))
        non_outlier_condition = self.get_non_outlier_condition(self.current_label, self.selected_element, old_id, new_id)

        if non_outlier_condition.sum() == 0:
            QMessageBox.information(self, "Info", "No non-outlier elements between selected row IDs.")
            return

        self.corrected_df.loc[non_outlier_condition, 'Corr Con'] *= ratio
        self.corrections_applied = True
        std_data = self.original_df[self.original_df['Type'] == 'Std'].copy(deep=True)
        updated_df = pd.concat([self.corrected_df, std_data], ignore_index=True)
        self.app.set_data(updated_df, for_results=True)
        self.app.notify_data_changed()
        self.display_non_outlier_ratios(self.current_label, self.selected_element)
        between_condition = self.get_non_outlier_condition(self.current_label, self.selected_element, old_id, new_id)
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)
        QMessageBox.information(self, "Result", f"Ratio correction ({ratio:.3f}) applied to {len(between_df)} non-outlier elements between row {self.selected_row_id_pair} for {self.selected_element} in {self.current_label}.")
        self.plot_trend(self.selected_element)

    def apply_all_corrections(self):
        """Apply corrections to all selected outliers."""
        selected_rows = []
        model = self.outliers_table.model()
        for row in range(model.rowCount()):
            if model.item(row, 0).checkState() == Qt.CheckState.Checked:
                selected_rows.append(row)

        if not selected_rows:
            QMessageBox.critical(self, "Error", "No outliers selected.")
            return

        corrections_applied = 0
        for row in selected_rows:
            label = model.data(model.index(row, 1))
            element = model.data(model.index(row, 2))
            for key, ratio in self.non_outlier_ratios.items():
                parts = key.split(':')
                if len(parts) != 3:
                    continue
                if parts[0] == label and parts[1] == element:
                    old_id, new_id = map(int, parts[2].split('->'))
                    non_outlier_condition = self.get_non_outlier_condition(label, element, old_id, new_id)
                    if non_outlier_condition.sum() > 0:
                        self.corrected_df.loc[non_outlier_condition, 'Corr Con'] *= ratio
                        corrections_applied += non_outlier_condition.sum()

        if corrections_applied > 0:
            self.corrections_applied = True
            std_data = self.original_df[self.original_df['Type'] == 'Std'].copy(deep=True)
            updated_df = pd.concat([self.corrected_df, std_data], ignore_index=True)
            self.app.set_data(updated_df, for_results=True)
            self.display_non_outlier_ratios(self.current_label, self.selected_element)
            between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
            between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
            self.display_between_df(between_df)
            QMessageBox.information(self, "Result", f"Applied corrections to {corrections_applied} non-outlier elements for selected outliers.")
            self.plot_trend(self.selected_element)
        else:
            QMessageBox.information(self, "Info", "No non-outlier elements found for selected outliers.")

    def mark_as_outlier(self):
        """Mark a non-outlier point as an outlier."""
        selection = self.non_outlier_table.selectionModel().selectedRows()
        if not selection or not self.current_label or not self.selected_element:
            QMessageBox.critical(self, "Error", "Select a non-outlier point first.")
            return

        model = self.non_outlier_table.model()
        row = selection[0].row()
        point_display = model.data(model.index(row, 2))
        if not point_display.startswith("Non-Outlier Point"):
            QMessageBox.critical(self, "Error", "Invalid non-outlier point selected.")
            return

        row_id = int(point_display.replace("Non-Outlier Point ", "").split('->')[0])
        user_key = f"{self.current_label}:{self.selected_element}"
        if user_key not in self.user_corrections:
            self.user_corrections[user_key] = {'outliers': [], 'non_outliers': []}
        if row_id not in self.user_corrections[user_key]['outliers']:
            self.user_corrections[user_key]['outliers'].append(row_id)
            self.user_corrections[user_key]['non_outliers'] = [x for x in self.user_corrections[user_key]['non_outliers'] if x != row_id]
            self.save_user_corrections()
            QMessageBox.information(self, "Info", f"Marked row {row_id} as outlier for {user_key}.")
            self.update_outlier_status()

    def mark_as_non_outlier(self):
        """Mark an outlier point as a non-outlier."""
        selection = self.ratios_table.selectionModel().selectedRows()
        if not selection or not self.current_label or not self.selected_element:
            QMessageBox.critical(self, "Error", "Select an outlier point first.")
            return

        model = self.ratios_table.model()
        row = selection[0].row()
        point_display = model.data(model.index(row, 2))
        if not point_display.startswith("Outlier Point"):
            QMessageBox.critical(self, "Error", "Invalid outlier point selected.")
            return

        row_id = int(point_display.replace("Outlier Point ", ""))
        user_key = f"{self.current_label}:{self.selected_element}"
        if user_key not in self.user_corrections:
            self.user_corrections[user_key] = {'outliers': [], 'non_outliers': []}
        if row_id not in self.user_corrections[user_key]['non_outliers']:
            self.user_corrections[user_key]['non_outliers'].append(row_id)
            self.user_corrections[user_key]['outliers'] = [x for x in self.user_corrections[user_key]['outliers'] if x != row_id]
            self.save_user_corrections()
            QMessageBox.information(self, "Info", f"Marked row {row_id} as non-outlier for {user_key}.")
            self.update_outlier_status()

    def skip_selected_outlier(self):
        """Skip a selected outlier."""
        selection = self.ratios_table.selectionModel().selectedRows()
        if not selection or not self.current_label or not self.selected_element:
            QMessageBox.critical(self, "Error", "Select an outlier point from the Outlier Points table.")
            return

        model = self.ratios_table.model()
        row = selection[0].row()
        point_display = model.data(model.index(row, 2))
        if not point_display.startswith("Outlier Point"):
            QMessageBox.critical(self, "Error", "Invalid outlier point selected.")
            return

        row_id = int(point_display.replace("Outlier Point ", ""))
        if self.current_label not in self.ignored_outliers:
            self.ignored_outliers[self.current_label] = set()
        self.ignored_outliers[self.current_label].add(row_id)
        QMessageBox.information(self, "Info", f"Ignored outlier at row {row_id} for {self.current_label}:{self.selected_element}.")
        self.update_outlier_status()

    def update_outlier_status(self):
        """Update outlier status and refresh tables."""
        if not self.current_label or not self.selected_element:
            return

        label_df = self.rm_df[self.rm_df['Solution Label'] == self.current_label].sort_values('row_id')
        if label_df.empty:
            return

        row_ids = label_df['row_id'].values
        values = pd.to_numeric(label_df[self.selected_element], errors='coerce').values
        valid_mask = ~np.isnan(values)
        if np.sum(valid_mask) < 2:
            return

        valid_values = values[valid_mask]
        valid_row_ids = row_ids[valid_mask]

        user_key = f"{self.current_label}:{self.selected_element}"
        user_outliers = set(self.user_corrections.get(user_key, {}).get('outliers', []))
        user_non_outliers = set(self.user_corrections.get(user_key, {}).get('non_outliers', []))
        ignored_outliers = set(self.ignored_outliers.get(self.current_label, set()))

        coefficients = np.polyfit(valid_row_ids, valid_values, 1)
        trendline = np.polyval(coefficients, valid_row_ids)
        residuals = valid_values - trendline
        residual_std = np.std(residuals) if len(residuals) > 0 else 1.0
        outlier_mask = np.abs(residuals) > float(self.threshold_entry.text()) * residual_std
        outlier_mask |= np.isin(valid_row_ids, list(user_outliers))
        outlier_mask &= ~np.isin(valid_row_ids, list(user_non_outliers))
        outlier_mask |= np.isin(valid_row_ids, list(ignored_outliers))

        self.outliers[self.current_label][self.selected_element] = int(np.sum(outlier_mask))
        self.ratios = {k: v for k, v in self.ratios.items() if not k.startswith(f"{self.current_label}:{self.selected_element}:")}
        self.non_outlier_ratios = {k: v for k, v in self.non_outlier_ratios.items() if not k.startswith(f"{self.current_label}:{self.selected_element}:")}

        outlier_row_ids = valid_row_ids[outlier_mask]
        for oid in outlier_row_ids:
            key = f"{self.current_label}:{self.selected_element}:{oid}"
            self.ratios[key] = np.nan

        non_outlier_values = valid_values[~outlier_mask]
        non_outlier_row_ids = valid_row_ids[~outlier_mask]
        if len(non_outlier_values) >= 2:
            for i in range(1, len(non_outlier_values)):
                old_val = non_outlier_values[i-1]
                new_val = non_outlier_values[i]
                if old_val == 0 or np.isnan(old_val) or np.isnan(new_val):
                    continue
                ratio = old_val / new_val
                key = f"{self.current_label}:{self.selected_element}:{non_outlier_row_ids[i-1]}->{non_outlier_row_ids[i]}"
                self.non_outlier_ratios[key] = ratio

        outlier_results = []
        for label in self.outliers:
            for element, count in self.outliers[label].items():
                if count > 0:
                    outlier_results.append({
                        'Solution Label': label,
                        'Element': element,
                        'Outliers Count': count
                    })
        if outlier_results:
            results_df = pd.DataFrame(outlier_results)
            self.display_outliers(results_df)
        else:
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(["Select", "Solution Label", "Element", "Outliers Count"])
            self.outliers_table.setModel(model)
            self.selected_outliers.clear()

        self.apply_corrections_for_label(self.current_label, self.selected_element)
        self.display_ratios(self.current_label, self.selected_element)
        self.display_non_outlier_ratios(self.current_label, self.selected_element)
        self.plot_trend(self.selected_element)
        between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)

    def plot_trend(self, element):
        """Placeholder for plotting trend."""
        if self.current_label is None:
            QMessageBox.critical(self, "Error", "No label selected.")
            return

        # Clear existing plot
        for widget in self.plot_frame.findChildren(QWidget):
            widget.deleteLater()

        placeholder_label = QLabel(f"Trend Plot for {element} (Not Implemented)")
        placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.plot_layout.addWidget(placeholder_label)