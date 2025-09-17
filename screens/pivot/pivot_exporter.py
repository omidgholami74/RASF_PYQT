import os
import platform
import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font as OpenPyXLFont, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class PivotExporter:
    """Handles exporting the pivot table to an Excel file."""
    def __init__(self, pivot_tab):
        self.pivot_tab = pivot_tab
        self.logger = pivot_tab.logger

    def export_pivot(self):
        """Export the pivot table to an Excel file with formatting."""
        if self.pivot_tab.pivot_data is None or self.pivot_tab.pivot_data.empty:
            QMessageBox.warning(self.pivot_tab, "Warning", "No data to export!")
            return

        try:
            filtered = self.pivot_tab.pivot_data.copy()
            for field, values in self.pivot_tab.row_filter_values.items():
                if field in filtered.columns:
                    selected = [k for k, v in values.items() if v]
                    if selected:
                        filtered = filtered[filtered[field].isin(selected)]

            selected_cols = ['Solution Label']
            for field, values in self.pivot_tab.column_filter_values.items():
                if field == 'Element':
                    selected_cols.extend([k for k, v in values.items() if v and k in filtered.columns])
            if len(selected_cols) > 1:
                filtered = filtered[selected_cols]

            filtered['Solution Label'] = pd.Categorical(filtered['Solution Label'], categories=self.pivot_tab.solution_label_order, ordered=True)
            filtered = filtered.sort_values('Solution Label').reset_index(drop=True)

            file_path = QFileDialog.getSaveFileName(self.pivot_tab, "Save Pivot Table", "", "Excel files (*.xlsx)")[0]
            if not file_path:
                return

            wb = Workbook()
            ws = wb.active
            ws.title = "Pivot Table"

            header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            first_col_fill = PatternFill(start_color="FFF5E4", end_color="FFF5E4", fill_type="solid")
            odd_fill = PatternFill(start_color="f5f5f5", end_color="f5f5f5", fill_type="solid")
            even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            diff_in_range_fill = PatternFill(start_color="ECFFC4", end_color="ECFFC4", fill_type="solid")
            diff_out_range_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            header_font = OpenPyXLFont(name="Segoe UI", size=12, bold=True)
            cell_font = OpenPyXLFont(name="Segoe UI", size=12)
            cell_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

            headers = list(filtered.columns)
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill = header_fill
                c.font = header_font
                c.alignment = cell_align
                c.border = thin_border
                ws.column_dimensions[get_column_letter(ci)].width = 15

            dec = int(self.pivot_tab.decimal_places.currentText())
            try:
                min_diff = float(self.pivot_tab.diff_min.text())
                max_diff = float(self.pivot_tab.diff_max.text())
            except ValueError:
                min_diff, max_diff = -12, 12

            row_idx = 2
            for _, row in filtered.iterrows():
                sol_label = row['Solution Label']
                for ci, val in enumerate(row, 1):
                    cell = ws.cell(row=row_idx, column=ci)
                    try:
                        cell.value = round(float(val), dec)
                        cell.number_format = f"0.{''.join(['0']*dec)}" if dec > 0 else "0"
                    except (ValueError, TypeError):
                        cell.value = "" if pd.isna(val) else str(val)
                    cell.font = cell_font
                    cell.alignment = cell_align
                    cell.border = thin_border
                    cell.fill = first_col_fill if ci == 1 else (even_fill if (row_idx - 1) % 2 == 0 else odd_fill)
                row_idx += 1

                if sol_label in self.pivot_tab._inline_crm_rows_display:
                    for crm_row_data, tags in self.pivot_tab._inline_crm_rows_display[sol_label]:
                        for ci, val in enumerate(crm_row_data, 1):
                            cell = ws.cell(row=row_idx, column=ci)
                            cell.value = str(val) if val else ""
                            cell.font = cell_font
                            cell.alignment = cell_align
                            cell.border = thin_border
                            if crm_row_data[0].endswith("CRM"):
                                cell.fill = first_col_fill if ci == 1 else PatternFill(start_color="FFF5E4", end_color="FFF5E4", fill_type="solid")
                            elif crm_row_data[0].endswith("Diff (%)"):
                                if tags[ci-1] == "in_range":
                                    cell.fill = diff_in_range_fill
                                elif tags[ci-1] == "out_range":
                                    cell.fill = diff_out_range_fill
                                else:
                                    cell.fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
                        row_idx += 1

            wb.save(file_path)
            QMessageBox.information(self.pivot_tab, "Success", "Pivot table exported successfully!")
            if QMessageBox.question(self.pivot_tab, "Open File", "Open the saved Excel file?") == QMessageBox.StandardButton.Yes:
                try:
                    if platform.system() == "Windows":
                        os.startfile(file_path)
                    elif platform.system() == "Darwin":
                        os.system(f"open '{file_path}'")
                    else:
                        os.system(f"xdg-open '{file_path}'")
                except Exception as e:
                    self.logger.error(f"Failed to open file: {str(e)}")
                    QMessageBox.warning(self.pivot_tab, "Error", f"Failed to open file: {str(e)}")

        except Exception as e:
            self.logger.error(f"Failed to export: {str(e)}")
            QMessageBox.warning(self.pivot_tab, "Error", f"Failed to export: {str(e)}")