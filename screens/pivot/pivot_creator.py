import re
import pandas as pd
from PyQt6.QtWidgets import QMessageBox
from .oxide_factors import oxide_factors

class PivotCreator:
    """Handles pivot table creation for the PivotTab."""
    def __init__(self, pivot_tab):
        self.pivot_tab = pivot_tab
        self.logger = pivot_tab.logger

    def create_pivot(self):
        """Create and populate the pivot table from the application data."""
        df = self.pivot_tab.app.get_data()
        if df is None or df.empty:
            QMessageBox.warning(self.pivot_tab, "Warning", "No data to display!")
            return

        try:
            self.pivot_tab.original_df = df.copy()
            df_filtered = df[df['Type'].isin(['Samp', 'Sample'])].copy()
            df_filtered['original_index'] = df_filtered.index
            grp = (df_filtered['Element'] != df_filtered['Element'].shift()).cumsum()
            df_filtered['Element'] = df_filtered['Element'].str.split('_').str[0]
            df_filtered['unique_id'] = df_filtered.groupby(['Solution Label', 'Element']).cumcount()

            def clean_label(label):
                m = re.search(r'(\d+)', str(label).replace(' ', ''))
                if m:
                    return f"{label.split()[0]} {m.group(1)}"
                return label

            self.pivot_tab.solution_label_order = sorted(df_filtered['Solution Label'].drop_duplicates().apply(clean_label).unique().tolist())
            self.pivot_tab.element_order = df_filtered['Element'].drop_duplicates().tolist()
            self.pivot_tab.element_selector.clear()
            self.pivot_tab.element_selector.addItems([""] + self.pivot_tab.element_order)

            value_column = 'Int' if self.pivot_tab.use_int_var.isChecked() else 'Corr Con'
            if value_column not in df_filtered.columns:
                QMessageBox.warning(self.pivot_tab, "Error", f"Column '{value_column}' not found in data!")
                return

            pivot_df = df_filtered.pivot_table(
                index=['Solution Label', 'unique_id'],
                columns='Element',
                values=value_column,
                aggfunc='first'
            ).reset_index()
            pivot_df = pivot_df.merge(
                df_filtered[['original_index', 'Solution Label', 'unique_id']],
                on=['Solution Label', 'unique_id'],
                how='left'
            ).sort_values('original_index').drop(columns=['original_index', 'unique_id']).drop_duplicates()
            
            if self.pivot_tab.use_oxide_var.isChecked():
                for col in pivot_df.columns:
                    if col != 'Solution Label' and col in oxide_factors:
                        _, factor = oxide_factors[col]
                        pivot_df[col] = pd.to_numeric(pivot_df[col], errors='coerce') * factor
            
            self.pivot_tab.pivot_data = pivot_df
            self.pivot_tab.column_widths.clear()
            self.pivot_tab.cached_formatted.clear()
            self.pivot_tab._inline_crm_rows.clear()
            self.pivot_tab._inline_crm_rows_display.clear()
            self.pivot_tab.row_filter_values.clear()
            self.pivot_tab.column_filter_values.clear()
            self.pivot_tab.update_pivot_display()

        except Exception as e:
            self.logger.error(f"Failed to create pivot table: {str(e)}")
            QMessageBox.warning(self.pivot_tab, "Pivot Error", f"Failed to create pivot table: {str(e)}")