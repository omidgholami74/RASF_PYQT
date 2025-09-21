from PyQt6.QtWidgets import (QWidget, QFrame, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QTableView, QHeaderView, QMessageBox, QMainWindow)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QStandardItemModel, QStandardItem
import pyqtgraph as pg
import pandas as pd
import numpy as np
import json
import os
import time
import logging
import uuid

# Enable antialiasing globally for pyqtgraph
pg.setConfigOptions(antialias=True)

# Setup logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class PlotWindow(QMainWindow):
    """A window to display a scatter plot for a selected column."""
    def __init__(self, title, df, x_column, y_column):
        super().__init__()
        self.setWindowTitle(title)
        self.setGeometry(150, 150, 800, 600)
        self.df = df
        self.x_column = x_column
        self.y_column = y_column
        self.setup_plot()
        logger.debug(f"PlotWindow created for {y_column}")

    def setup_plot(self):
        """Setup the scatter plot in the window."""
        logger.debug(f"Setting up scatter plot for {self.y_column} vs {self.x_column}")

        valid_data = 0
        x_values = []
        y_values = []
        for _, row in self.df.iterrows():
            x_value = row[self.x_column]
            y_value = row[self.y_column]
            if pd.notna(x_value) and pd.notna(y_value):
                try:
                    x_float = float(x_value)
                    y_float = float(y_value)
                    x_values.append(x_float)
                    y_values.append(y_float)
                    valid_data += 1
                except (ValueError, TypeError) as e:
                    logger.debug(f"Skipping invalid data point ({x_value}, {y_value}) for {self.y_column}: {e}")
                    continue

        if valid_data == 0:
            logger.warning(f"No valid data to plot for {self.y_column}")
            self.setCentralWidget(QLabel("No valid data to plot"))
            return

        # Use pyqtgraph
        layout_widget = pg.GraphicsLayoutWidget()
        layout_widget.setBackground('w')  # Set background to white
        title_item = pg.LabelItem(f"Scatter Plot of {self.y_column} vs {self.x_column}", size='12pt', bold=True)
        layout_widget.addItem(title_item, row=0, col=0)
        plot = layout_widget.addPlot(row=1, col=0)
        plot.plot(x_values, y_values, pen=None, symbol='o', symbolSize=8, symbolBrush='b', name=f"{self.y_column} Data")
        plot.setLabel('bottom', self.x_column)
        plot.setLabel('left', self.y_column)

        x_values = pd.to_numeric(self.df[self.x_column], errors='coerce').dropna()
        if not x_values.empty:
            plot.setXRange(x_values.min(), x_values.max())

        y_values = pd.to_numeric(self.df[self.y_column], errors='coerce').dropna()
        if not y_values.empty:
            plot.setYRange(y_values.min(), y_values.max())

        legend = plot.addLegend(offset=(10, 10))
        self.setCentralWidget(layout_widget)
        logger.debug(f"Scatter plot setup completed for {self.y_column}")

class CheckRMFrame(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.reset_state()
        self.corrections_file = "user_corrections.json"
        self.load_user_corrections()
        self.setup_ui()
        logger.debug("CheckRMFrame initialized")

    def reset_state(self):
        """Reset all state variables to ensure a clean slate."""
        self.rm_df = None
        self.positions_df = None
        self.original_df = None
        self.corrected_df = None
        self.outliers = {}
        self.ratios = {}
        self.non_outlier_ratios = {}
        self.selected_outliers = {}
        self.ignored_outliers = {}
        self.current_label = None
        self.selected_element = None
        self.selected_row_id_pair = None
        self.corrections_applied = False
        self.current_between_df = None
        self.plot_windows = {}
        self.user_corrections = {}
        logger.debug("State reset")

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
        logger.debug("Styles configured")

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

        threshold_label = QLabel("Threshold (%):")
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

        # Buttons for plotting columns
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.addWidget(QLabel("Plot Columns:"))
        numeric_cols = ["Previous Corr Con", "Corr Con", "Soln Conc", "Intense", "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)"]
        self.column_buttons = {}
        for col in numeric_cols:
            btn = QPushButton(f"Plot {col}")
            btn.clicked.connect(lambda checked, c=col: self.open_plot_window(c))
            button_layout.addWidget(btn)
            self.column_buttons[col] = btn
            logger.debug(f"Connected button for {col}")
        right_layout.addWidget(button_frame)

        self.corrected_table = QTableView()
        self.corrected_table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.corrected_table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.corrected_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.corrected_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.corrected_table)

        content_layout.addWidget(right_frame, stretch=1)
        main_layout.addWidget(content_frame)
        logger.debug("UI setup completed")

    def open_plot_window(self, column):
        """Open a new scatter plot window for the selected column using data from current_between_df."""
        logger.debug(f"open_plot_window called for column: {column}")
        logger.debug(f"current_between_df is None: {self.current_between_df is None}")

        if self.current_between_df is None or self.current_between_df.empty:
            logger.debug("current_between_df is None or empty")
            QMessageBox.warning(self, "Warning", "No data available to plot.")
            return

        plot_df = self.current_between_df.copy()
        logger.debug(f"Using current_between_df with {len(plot_df)} rows")
        logger.debug(f"current_between_df columns: {plot_df.columns.tolist()}")

        try:
            if column not in plot_df.columns:
                logger.debug(f"Column {column} not in plot_df, adding with NaN")
                plot_df[column] = np.nan

            plot_df = plot_df.reset_index(drop=True)
            plot_df['index'] = plot_df.index

            logger.debug(f"plot_df before filtering: {len(plot_df)} rows")
            logger.debug(f"NaN count in {column}: {plot_df[column].isna().sum()}")
            logger.debug(f"NaN count in index: {plot_df['index'].isna().sum()}")

            plot_df = plot_df[['index', column]].copy()
            plot_df['index'] = pd.to_numeric(plot_df['index'], errors='coerce')
            plot_df[column] = pd.to_numeric(plot_df[column], errors='coerce')
            valid_data = plot_df.dropna()

            skipped_count = len(plot_df) - len(valid_data)
            if skipped_count > 0:
                logger.debug(f"Skipped {skipped_count} invalid or NaN data points for column {column}")

            if valid_data.empty:
                logger.debug(f"No valid data to plot for column {column}")
                QMessageBox.warning(self, "Warning", f"No valid data to plot for {column}.")
                return

            logger.debug(f"Valid data for plotting: {len(valid_data)} rows")
            logger.debug(f"Valid data sample: {valid_data.head().to_dict()}")

            window = PlotWindow(f"Scatter Plot for {column}", valid_data, "index", column)
            window.show()
            window.raise_()
            window.activateWindow()
            window_id = str(uuid.uuid4())
            self.plot_windows[window_id] = window
            logger.debug(f"Plot window opened for {column} with ID {window_id}")
        except Exception as e:
            logger.error(f"Error creating plot window for {column}: {e}")
            QMessageBox.critical(self, "Error", f"Error plotting {column}: {str(e)}")
            return

    def check_rm_changes(self):
        """Check for RM changes and detect outliers."""
        start_time = time.time()
        self.reset_state()  # Reset state to ensure a clean slate

        try:
            threshold_percent = float(self.threshold_entry.text())
            threshold = threshold_percent / 100
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter a valid number for threshold (%).")
            return

        keyword = self.keyword_entry.text().strip()
        if not keyword:
            QMessageBox.critical(self, "Error", "Please enter a valid keyword.")
            return

        df = self.app.get_data()
        if df is None or df.empty:
            QMessageBox.critical(self, "Error", "No data loaded.")
            return

        required_columns = ['Solution Label', 'Element', 'Type', 'Corr Con', 'Act Wgt', 'Act Vol', 'Coeff 1', 'Coeff 2']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            QMessageBox.critical(self, "Error", f"Missing required columns: {missing_columns}")
            return

        # Log input DataFrame for debugging
        logger.debug(f"Input DataFrame shape: {df.shape}")
        logger.debug(f"Input DataFrame columns: {df.columns.tolist()}")
        logger.debug(f"Input DataFrame index duplicates: {df.index.duplicated().sum()}")

        # Initialize self.original_df
        self.original_df = df.copy(deep=True)
        if 'original_index' in self.original_df.columns:
            logger.debug("Dropping existing 'original_index' from original_df")
            self.original_df = self.original_df.drop(columns=['original_index'])
        if 'row_id' in self.original_df.columns:
            logger.debug("Dropping existing 'row_id' from original_df")
            self.original_df = self.original_df.drop(columns=['row_id'])
        self.original_df = self.original_df.reset_index(drop=True)
        self.original_df['original_index'] = self.original_df.index
        logger.debug(f"original_df columns: {self.original_df.columns.tolist()}")
        logger.debug(f"original_index duplicates in original_df: {self.original_df['original_index'].duplicated().sum()}")

        # Filter for 'Samp' type
        df_filtered = df[df['Type'] == 'Samp'].copy(deep=True)
        if df_filtered.empty:
            QMessageBox.critical(self, "Error", f"No data with Type='Samp' found. Number of rows: {len(df)}")
            return
        if 'original_index' in df_filtered.columns:
            logger.debug("Dropping existing 'original_index' from df_filtered")
            df_filtered = df_filtered.drop(columns=['original_index'])
        if 'row_id' in df_filtered.columns:
            logger.debug("Dropping existing 'row_id' from df_filtered")
            df_filtered = df_filtered.drop(columns=['row_id'])
        df_filtered = df_filtered.reset_index(drop=True)
        df_filtered['original_index'] = df_filtered.index
        logger.debug(f"df_filtered columns: {df_filtered.columns.tolist()}")
        logger.debug(f"original_index duplicates in df_filtered: {df_filtered['original_index'].duplicated().sum()}")

        # Clean 'Solution Label' and add 'row_id'
        df_filtered['Solution Label'] = df_filtered['Solution Label'].str.replace(
            rf'^{keyword}\s*[-]?\s*(\d*)$', rf'{keyword}\1', regex=True
        )
        df_filtered['row_id'] = df_filtered.groupby(['Solution Label', 'Element']).cumcount()

        # Log DataFrames before merge
        logger.debug(f"self.original_df head:\n{self.original_df.head().to_string()}")
        logger.debug(f"df_filtered head:\n{df_filtered.head().to_string()}")

        # Perform the merge
        try:
            self.original_df = self.original_df.merge(
                df_filtered[['original_index', 'Solution Label', 'Element', 'row_id']],
                on=['Solution Label', 'Element', 'original_index'],
                how='left'
            )
            logger.debug(f"Post-merge original_df columns: {self.original_df.columns.tolist()}")
            logger.debug(f"Post-merge original_df shape: {self.original_df.shape}")
            logger.debug(f"Post-merge row_id NaN count: {self.original_df['row_id'].isna().sum()}")
        except Exception as e:
            logger.error(f"Merge failed: {e}")
            QMessageBox.critical(self, "Error", f"Merge failed: {str(e)}")
            return

        # Check if 'row_id' exists after merge
        if 'row_id' not in self.original_df.columns:
            logger.error("Merge did not include 'row_id' column")
            QMessageBox.critical(self, "Error", "Merge failed to include 'row_id' column. Check data consistency.")
            return

        # Handle 'row_id' column
        self.original_df['row_id'] = self.original_df['row_id'].fillna(-1).astype(int)

        # Convert numeric columns
        numeric_columns = ['Corr Con', 'Act Wgt', 'Act Vol', 'Coeff 1', 'Coeff 2']
        for col in numeric_columns:
            df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce')
            self.original_df[col] = pd.to_numeric(self.original_df[col], errors='coerce')
            logger.debug(f"NaN count in {col} (original_df): {self.original_df[col].isna().sum()}")
            logger.debug(f"NaN count in {col} (df_filtered): {df_filtered[col].isna().sum()}")

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
                outlier_mask = np.abs(residuals / (trendline + 1e-10)) > threshold
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
        logger.debug("Outliers table displayed")

    def toggle_checkbox(self, index):
        """Toggle checkbox state in the outliers table."""
        if index.column() == 0:
            model = self.outliers_table.model()
            item = model.item(index.row(), 0)
            if item.isCheckable():
                new_state = Qt.CheckState.Checked if item.checkState() == Qt.CheckState.Unchecked else Qt.CheckState.Unchecked
                item.setCheckState(new_state)
                item.setText("☑" if new_state == Qt.CheckState.Checked else "☐")
                logger.debug(f"Checkbox toggled at row {index.row()} to state {new_state}")

    def apply_corrections_for_label(self, label, element):
        """Apply corrections for a specific label and element to achieve zero slope."""
        try:
            threshold_percent = float(self.threshold_entry.text())
            threshold = threshold_percent / 100
        except ValueError:
            threshold = 0.02

        label_df = self.rm_df[self.rm_df['Solution Label'] == label].sort_values('row_id')
        if label_df.empty:
            logger.warning(f"No data for label {label}")
            return

        row_ids = label_df['row_id'].values
        values = pd.to_numeric(label_df[element], errors='coerce').values
        valid_mask = ~np.isnan(values)
        if np.sum(valid_mask) < 2:
            logger.warning(f"Insufficient valid data for {element} in {label}")
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
        outlier_mask = np.abs(residuals / (trendline + 1e-10)) > threshold
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
            logger.debug(f"Applied mean correction ({mean_value:.3f}) to {np.sum(condition & valid_rows)} rows for {label}:{element}")

            std_data = self.original_df[self.original_df['Type'] == 'Std'].copy(deep=True)
            updated_df = pd.concat([self.corrected_df, std_data], ignore_index=True)
            self.app.set_data(updated_df, for_results=True)
            self.app.notify_data_changed()

    def get_non_outlier_condition(self, label, element, old_id, new_id):
        """Get condition for non-outlier rows between old_id and new_id."""
        max_old = self.positions_df[(self.positions_df['Solution Label'] == label) & (self.positions_df['row_id'] == old_id)]['max'].values
        min_new = self.positions_df[(self.positions_df['Solution Label'] == label) & (self.positions_df['row_id'] == new_id)]['min'].values
        if len(max_old) == 0 or len(min_new) == 0:
            logger.warning(f"No valid positions found for {label} between row_ids {old_id} and {new_id}")
            return pd.Series(False, index=self.corrected_df.index)
        condition = (self.corrected_df['original_index'] > max_old[0]) & (self.corrected_df['original_index'] < min_new[0])
        logger.debug(f"Non-outlier condition for {label}:{element} between {old_id}->{new_id}: {condition.sum()} rows")
        return condition

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
        logger.debug(f"No non-outlier condition found for {label}:{element}")
        return pd.Series(False, index=self.corrected_df.index)

    def display_between_df(self, df):
        """Display corrected DataFrame in the corrected_table and store computed values in current_between_df."""
        logger.debug("Entering display_between_df")
        self.current_between_df = df.copy() if df is not None and not df.empty else pd.DataFrame(columns=["Solution Label", "Element", "Previous Corr Con", "Corr Con", "Soln Conc", "Intense", "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)", "row_id", "original_index", "Act Wgt", "Act Vol", "Coeff 1", "Coeff 2"])
        logger.debug(f"current_between_df initialized with {len(self.current_between_df)} rows")
        logger.debug(f"current_between_df columns: {self.current_between_df.columns.tolist()}")

        required_columns = ["Solution Label", "Element", "Previous Corr Con", "Corr Con", "Soln Conc", "Intense", "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)", "row_id", "original_index", "Act Wgt", "Act Vol", "Coeff 1", "Coeff 2"]
        for col in required_columns:
            if col not in self.current_between_df.columns:
                self.current_between_df[col] = np.nan
                logger.debug(f"Added missing column {col} with NaN")

        self.current_between_df["Previous Corr Con"] = np.nan
        self.current_between_df["Soln Conc"] = np.nan
        self.current_between_df["Intense"] = np.nan
        self.current_between_df["Previous Intense"] = np.nan
        self.current_between_df["Soln Conc (No RM)"] = np.nan
        self.current_between_df["Intense (No RM)"] = np.nan

        for idx, row in self.current_between_df.iterrows():
            original_row = self.original_df[
                (self.original_df['Solution Label'] == row['Solution Label']) &
                (self.original_df['Element'] == row['Element']) &
                (self.original_df['original_index'] == row['original_index'])
            ]
            prev_corr_con = np.nan
            if not original_row.empty:
                prev_corr_con = pd.to_numeric(original_row['Corr Con'].iloc[0], errors='coerce')
                logger.debug(f"Found matching original row for {row['Solution Label']}:{row['Element']}:{row['original_index']}, Prev Corr Con: {prev_corr_con}")
            else:
                logger.debug(f"No matching original row for {row['Solution Label']}:{row['Element']}:{row['original_index']}")

            try:
                corr_con = pd.to_numeric(row['Corr Con'], errors='coerce') if pd.notna(row['Corr Con']) else np.nan
                act_wgt = pd.to_numeric(row['Act Wgt'], errors='coerce') if pd.notna(row['Act Wgt']) else np.nan
                act_vol = pd.to_numeric(row['Act Vol'], errors='coerce') if pd.notna(row['Act Vol']) and row['Act Vol'] != 0 else np.nan
                coeff_1 = pd.to_numeric(row['Coeff 1'], errors='coerce') if pd.notna(row['Coeff 1']) else np.nan
                coeff_2 = pd.to_numeric(row['Coeff 2'], errors='coerce') if pd.notna(row['Coeff 2']) else np.nan

                logger.debug(f"Row {idx} inputs: Corr Con={corr_con}, Act Wgt={act_wgt}, Act Vol={act_vol}, Coeff 1={coeff_1}, Coeff 2={coeff_2}")

                soln_conc = corr_con * act_wgt / act_vol if pd.notna(corr_con) and pd.notna(act_wgt) and pd.notna(act_vol) else np.nan
                intense = soln_conc * coeff_2 + coeff_1 if pd.notna(soln_conc) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan
                soln_conc_no_rm = prev_corr_con * act_wgt / act_vol if pd.notna(prev_corr_con) and pd.notna(act_wgt) and pd.notna(act_vol) else np.nan
                intense_no_rm = soln_conc_no_rm * coeff_2 + coeff_1 if pd.notna(soln_conc_no_rm) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan
                prev_intense = soln_conc_no_rm * coeff_2 + coeff_1 if pd.notna(soln_conc_no_rm) and pd.notna(coeff_2) and pd.notna(coeff_1) else np.nan

                self.current_between_df.at[idx, "Previous Corr Con"] = prev_corr_con
                self.current_between_df.at[idx, "Soln Conc"] = soln_conc
                self.current_between_df.at[idx, "Intense"] = intense
                self.current_between_df.at[idx, "Previous Intense"] = prev_intense
                self.current_between_df.at[idx, "Soln Conc (No RM)"] = soln_conc_no_rm
                self.current_between_df.at[idx, "Intense (No RM)"] = intense_no_rm

                logger.debug(f"Computed for row {idx}: Soln Conc={soln_conc}, Intense={intense}, Previous Corr Con={prev_corr_con}, Previous Intense={prev_intense}, Soln Conc (No RM)={soln_conc_no_rm}, Intense (No RM)={intense_no_rm}")
            except (TypeError, ValueError) as e:
                logger.error(f"Error computing values for row {idx}: {e}")

        numeric_cols = ["Previous Corr Con", "Corr Con", "Soln Conc", "Intense", "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)"]
        for col in numeric_cols:
            if col in self.current_between_df.columns:
                logger.debug(f"NaN count in {col}: {self.current_between_df[col].isna().sum()}")

        model = QStandardItemModel()
        columns = ["Solution Label", "Element", "Previous Corr Con", "Corr Con", "Soln Conc", "Intense", 
                   "Previous Intense", "Soln Conc (No RM)", "Intense (No RM)", "row_id"]
        model.setHorizontalHeaderLabels(columns)

        if self.current_between_df.empty:
            logger.debug("No data to display in Between Elements Preview table")
            self.corrected_table.setModel(model)
            return

        for _, row in self.current_between_df.iterrows():
            model.appendRow([
                QStandardItem(str(row['Solution Label'])),
                QStandardItem(str(row['Element'])),
                QStandardItem(f"{row['Previous Corr Con']:.3f}" if pd.notna(row['Previous Corr Con']) else "N/A"),
                QStandardItem(f"{row['Corr Con']:.3f}" if pd.notna(row['Corr Con']) else "N/A"),
                QStandardItem(f"{row['Soln Conc']:.3f}" if pd.notna(row['Soln Conc']) else "N/A"),
                QStandardItem(f"{row['Intense']:.3f}" if pd.notna(row['Intense']) else "N/A"),
                QStandardItem(f"{row['Previous Intense']:.3f}" if pd.notna(row['Previous Intense']) else "N/A"),
                QStandardItem(f"{row['Soln Conc (No RM)']:.3f}" if pd.notna(row['Soln Conc (No RM)']) else "N/A"),
                QStandardItem(f"{row['Intense (No RM)']:.3f}" if pd.notna(row['Intense (No RM)']) else "N/A"),
                QStandardItem(str(row['row_id']))
            ])

        self.corrected_table.setModel(model)
        self.corrected_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        logger.debug("display_between_df completed")

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
        logger.debug(f"Outlier selected: {self.current_label}:{self.selected_element}")

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
        logger.debug(f"Ratio selected: {self.current_label}:{self.selected_element}")

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
        logger.debug(f"Non-outlier selected: {self.current_label}:{self.selected_element}")

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
        self.apply_corrections_for_label(self.current_label, self.selected_element)
        self.display_non_outlier_ratios(self.current_label, self.selected_element)
        between_condition = self.get_non_outlier_condition(self.current_label, self.selected_element, old_id, new_id)
        between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
        self.display_between_df(between_df)
        self.plot_trend(self.selected_element)
        QMessageBox.information(self, "Result", f"Ratio correction ({ratio:.3f}) applied to {len(between_df)} non-outlier elements between row {self.selected_row_id_pair} for {self.selected_element} in {self.current_label}.")
        logger.debug(f"Ratio correction applied: {ratio:.3f} to {len(between_df)} rows")

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
            self.apply_corrections_for_label(self.current_label, self.selected_element)
            self.display_non_outlier_ratios(self.current_label, self.selected_element)
            between_condition = self.get_first_non_outlier_condition(self.current_label, self.selected_element)
            between_df = self.corrected_df[between_condition & (self.corrected_df['Element'] == self.selected_element)]
            self.display_between_df(between_df)
            self.plot_trend(self.selected_element)
            QMessageBox.information(self, "Result", f"Applied corrections to {corrections_applied} non-outlier elements for selected outliers.")
            logger.debug(f"Applied all corrections to {corrections_applied} rows")
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
            self.update_outlier_status()
            QMessageBox.information(self, "Info", f"Marked row {row_id} as outlier for {user_key}.")
            logger.debug(f"Marked row {row_id} as outlier for {user_key}")

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
            self.update_outlier_status()
            QMessageBox.information(self, "Info", f"Marked row {row_id} as non-outlier for {user_key}.")
            logger.debug(f"Marked row {row_id} as non-outlier for {user_key}")

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
        self.update_outlier_status()
        QMessageBox.information(self, "Info", f"Ignored outlier at row {row_id} for {self.current_label}:{self.selected_element}.")
        logger.debug(f"Ignored outlier at row {row_id} for {self.current_label}:{self.selected_element}")

    def update_outlier_status(self):
        """Update outlier status and refresh tables and plots."""
        try:
            threshold_percent = float(self.threshold_entry.text())
            threshold = threshold_percent / 100
        except ValueError:
            threshold = 0.02

        if not self.current_label or not self.selected_element:
            logger.warning("No current label or element selected for updating outlier status")
            return

        label_df = self.rm_df[self.rm_df['Solution Label'] == self.current_label].sort_values('row_id')
        if label_df.empty:
            logger.warning(f"No data for label {self.current_label}")
            return

        row_ids = label_df['row_id'].values
        values = pd.to_numeric(label_df[self.selected_element], errors='coerce').values
        valid_mask = ~np.isnan(values)
        if np.sum(valid_mask) < 2:
            logger.warning(f"Insufficient valid data for {self.selected_element} in {self.current_label}")
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
        outlier_mask = np.abs(residuals / (trendline + 1e-10)) > threshold
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
        logger.debug("Outlier status updated")

    def plot_trend(self, element):
        """Plot trend for the specified element in the current label."""
        if self.current_label is None:
            QMessageBox.critical(self, "Error", "No label selected.")
            return

        def render_plot():
            for widget in self.plot_frame.findChildren(QWidget):
                widget.deleteLater()

            try:
                threshold_percent = float(self.threshold_entry.text())
                threshold = threshold_percent / 100
            except ValueError:
                threshold = 0.02

            label_df_original = self.rm_df[self.rm_df['Solution Label'] == self.current_label].sort_values('row_id')
            if label_df_original.empty:
                QMessageBox.critical(self, "Error", f"No data found for {self.current_label}.")
                return

            label_df_original = label_df_original.reset_index(drop=True)
            sample_index = np.array(label_df_original.index)
            values_original = pd.to_numeric(label_df_original[element], errors='coerce').values
            valid_mask = ~np.isnan(values_original)
            if np.sum(valid_mask) < 2:
                QMessageBox.information(self, "Info", f"Insufficient valid data for {element} in {self.current_label}.")
                return

            valid_sample_index = sample_index[valid_mask]
            valid_values_original = values_original[valid_mask]

            corrected_condition = (self.corrected_df['Solution Label'] == self.current_label) & (self.corrected_df['Element'] == element)
            corrected_df_label = self.corrected_df[corrected_condition].sort_values('row_id')
            values_corrected = pd.to_numeric(corrected_df_label['Corr Con'], errors='coerce').values
            if len(values_corrected) != len(values_original):
                QMessageBox.warning(self, "Warning", "Mismatch in original and corrected data lengths.")
                return
            valid_values_corrected = values_corrected[valid_mask]

            user_key = f"{self.current_label}:{element}"
            user_outliers = set(self.user_corrections.get(user_key, {}).get('outliers', []))
            user_non_outliers = set(self.user_corrections.get(user_key, {}).get('non_outliers', []))
            ignored_outliers = set(self.ignored_outliers.get(self.current_label, set()))

            coefficients = np.polyfit(valid_sample_index, valid_values_original, 1)
            trendline = np.polyval(coefficients, valid_sample_index)
            residuals = valid_values_original - trendline
            outlier_mask = np.abs(residuals / (trendline + 1e-10)) > threshold
            outlier_mask |= np.isin(valid_sample_index, list(user_outliers))
            outlier_mask &= ~np.isin(valid_sample_index, list(user_non_outliers))
            outlier_mask |= np.isin(valid_sample_index, list(ignored_outliers))

            non_outlier_mask = ~outlier_mask
            valid_sample_index_no_outliers = valid_sample_index[non_outlier_mask]
            valid_values_original_no_outliers = valid_values_original[non_outlier_mask]
            valid_values_corrected_no_outliers = valid_values_corrected[non_outlier_mask]

            if len(valid_values_corrected_no_outliers) < 1:
                QMessageBox.information(self, "Info", f"No non-outlier data available for {element} in {self.current_label}.")
                return

            coefficients_before = np.polyfit(valid_sample_index, valid_values_original, 1)
            trendline_before = np.polyval(coefficients_before, valid_sample_index)

            if len(valid_sample_index_no_outliers) >= 2:
                coefficients_after = np.polyfit(valid_sample_index_no_outliers, valid_values_corrected_no_outliers, 1)
            else:
                mean_val = np.mean(valid_values_corrected_no_outliers) if len(valid_values_corrected_no_outliers) > 0 else 0
                coefficients_after = [0, mean_val]

            trendline_after = np.polyval(coefficients_after, valid_sample_index)

            layout_widget = pg.GraphicsLayoutWidget()
            layout_widget.setBackground('w')
            title_item = pg.LabelItem(f"Trend for {element} ({self.current_label})", size='12pt', bold=True)
            layout_widget.addItem(title_item, row=0, col=0)
            plot = layout_widget.addPlot(row=1, col=0)

            plot.plot(valid_sample_index_no_outliers, valid_values_original_no_outliers, name=f"Original {element}", pen=pg.mkPen('b', width=2))
            plot.plot(valid_sample_index[outlier_mask], valid_values_original[outlier_mask], pen=None, symbol='o', symbolSize=8, symbolBrush='r', name="Outlier Points")
            plot.plot(valid_sample_index, trendline_before, name=f"Trend Before (slope={coefficients_before[0]:.3f})", pen=pg.mkPen('b', style=Qt.PenStyle.DashLine))
            plot.plot(valid_sample_index_no_outliers, valid_values_corrected_no_outliers, name=f"Corrected {element}", pen=pg.mkPen('g', width=2))
            plot.plot(valid_sample_index, trendline_after, name=f"Trend After (slope={coefficients_after[0]:.3f})", pen=pg.mkPen('g', style=Qt.PenStyle.DashLine))

            plot.setLabel('bottom', "Index")
            plot.setLabel('left', element)
            y_min = min(np.min(valid_values_original), np.min(valid_values_corrected)) - 0.1 * (max(np.max(valid_values_original), np.max(valid_values_corrected)) - min(np.min(valid_values_original), np.min(valid_values_corrected)))
            y_max = max(np.max(valid_values_original), np.max(valid_values_corrected)) + 0.1 * (max(np.max(valid_values_original), np.max(valid_values_corrected)) - min(np.min(valid_values_original), np.min(valid_values_corrected)))
            plot.setXRange(valid_sample_index.min(), valid_sample_index.max())
            plot.setYRange(y_min, y_max)
            plot.addLegend(offset=(10, 10))

            self.plot_layout.addWidget(layout_widget)
            logger.debug(f"Trend plot rendered for {element} in {self.current_label}")

        QTimer.singleShot(0, render_plot)

    def display_ratios(self, label, element):
        """Display outlier points and their ratios in the ratios_table."""
        logger.debug(f"Displaying ratios for {label}:{element}")
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Solution Label", "Element", "Point", "Ratio"])

        for key, ratio in self.ratios.items():
            parts = key.split(':')
            if len(parts) != 3 or parts[0] != label or parts[1] != element:
                continue
            row_id = parts[2]
            model.appendRow([
                QStandardItem(str(label)),
                QStandardItem(str(element)),
                QStandardItem(f"Outlier Point {row_id}"),
                QStandardItem(f"{ratio:.3f}" if pd.notna(ratio) else "N/A")
            ])

        self.ratios_table.setModel(model)
        self.ratios_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        logger.debug(f"Ratios table populated with {model.rowCount()} rows for {label}:{element}")

    def display_non_outlier_ratios(self, label, element):
        """Display non-outlier points and their ratios in the non_outlier_table."""
        logger.debug(f"Displaying non-outlier ratios for {label}:{element}")
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Solution Label", "Element", "Point", "Ratio"])

        for key, ratio in self.non_outlier_ratios.items():
            parts = key.split(':')
            if len(parts) != 3 or parts[0] != label or parts[1] != element:
                continue
            point_display = parts[2]
            model.appendRow([
                QStandardItem(str(label)),
                QStandardItem(str(element)),
                QStandardItem(f"Non-Outlier Point {point_display}"),
                QStandardItem(f"{ratio:.3f}" if pd.notna(ratio) else "N/A")
            ])

        self.non_outlier_table.setModel(model)
        self.non_outlier_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        logger.debug(f"Non-outlier table populated with {model.rowCount()} rows for {label}:{element}")