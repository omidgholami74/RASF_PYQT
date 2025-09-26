import re
import pandas as pd
from PyQt6.QtWidgets import QCheckBox, QMessageBox, QDialog, QVBoxLayout, QRadioButton, QPushButton, QLabel
from .oxide_factors import oxide_factors

class CRMManager:
    """Manages CRM-related operations for the PivotTab."""
    def __init__(self, pivot_tab):
        self.pivot_tab = pivot_tab
        self.logger = pivot_tab.logger

    def check_rm(self):
        """Check Reference Materials (RM) against the CRM database and update inline CRM rows."""
        if self.pivot_tab.pivot_data is None or self.pivot_tab.pivot_data.empty:
            QMessageBox.warning(self.pivot_tab, "Warning", "No pivot data available!")
            self.logger.warning("No pivot data available in check_rm")
            return

        try:
            conn = self.pivot_tab.app.crm_tab.conn
            if conn is None:
                self.pivot_tab.app.crm_tab.init_db()
                conn = self.pivot_tab.app.crm_tab.conn
                if conn is None:
                    QMessageBox.warning(self.pivot_tab, "Error", "Failed to connect to CRM database!")
                    self.logger.error("Failed to connect to CRM database")
                    return

            # Filter rows with 'CRM', 'par', or 'OREAS' in Solution Label
            crm_rows = self.pivot_tab.pivot_data[
                self.pivot_tab.pivot_data['Solution Label'].str.contains('CRM|par|OREAS', case=False, na=False)
            ].copy()
            print(f"CRM Rows found: {crm_rows['Solution Label'].tolist()}")
            if crm_rows.empty:
                QMessageBox.information(self.pivot_tab, "Info", "No CRM, par, or OREAS rows found in pivot data!")
                self.logger.info("No CRM, par, or OREAS rows found in pivot data")
                return

            # Query the pivoted CRM table
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(pivot_crm)")
            cols = [x[1] for x in cursor.fetchall()]
            required = {'CRM ID'}
            if not required.issubset(cols):
                QMessageBox.warning(self.pivot_tab, "Error", "pivot_crm table missing required columns!")
                self.logger.error("pivot_crm table missing required columns")
                return

            # Debug: Print available CRM IDs and columns
            cursor.execute("SELECT DISTINCT [CRM ID] FROM pivot_crm")
            db_crm_ids = cursor.fetchall()
            print(f"Available CRM IDs in database: {db_crm_ids}")
            print(f"pivot_crm columns: {cols}")
            print(f"pivot_data columns: {list(self.pivot_tab.pivot_data.columns)}")

            # Create element-to-wavelength mapping
            element_to_columns = {}
            if self.pivot_tab.original_df is not None and 'Element' in self.pivot_tab.original_df.columns:
                # Use original_df to map elements to wavelength columns
                for col in self.pivot_tab.pivot_data.columns:
                    if col == 'Solution Label':
                        continue
                    element = col.split()[0].strip()  # e.g., 'Cu 324.754' -> 'Cu'
                    if element not in element_to_columns:
                        element_to_columns[element] = []
                    element_to_columns[element].append(col)
            else:
                # Fallback: Parse pivot_data column names
                for col in self.pivot_tab.pivot_data.columns:
                    if col == 'Solution Label':
                        continue
                    element = col.split()[0].strip()  # e.g., 'Cu 324.754' -> 'Cu'
                    if element not in element_to_columns:
                        element_to_columns[element] = []
                    element_to_columns[element].append(col)
            print(f"Element to columns mapping: {element_to_columns}")

            dec = int(self.pivot_tab.decimal_places.currentText())
            self.pivot_tab._inline_crm_rows.clear()
            self.pivot_tab.included_crms.clear()

            for _, row in crm_rows.iterrows():
                label = row['Solution Label']
                print(f"Processing label: {label}")
                # Regex: Match digits and optional letter, possibly preceded by CRM/OREAS or followed by par
                m = re.search(r'(?i)(?:CRM|OREAS)?\s*(\d+[a-zA-Z]?)(?:\s*par)?', str(label))
                if not m:
                    self.logger.warning(f"No valid CRM ID found in label: {label}")
                    continue
                crm_id_part = m.group(1)
                crm_id_string = f"OREAS {crm_id_part}"
                print(f"Querying CRM ID: {crm_id_string}")

                # Query pivot_crm for matching CRM ID
                cursor.execute(
                    "SELECT * FROM pivot_crm WHERE [CRM ID] LIKE ?",
                    (f"OREAS {crm_id_part}%",)
                )
                crm_data = cursor.fetchall()
                print(f"Partial match results for OREAS {crm_id_part}%: {crm_data}")
                if not crm_data:
                    self.logger.warning(f"No CRM data found for {crm_id_string} or partial matches")
                    continue

                # Get column names for the pivoted data
                cursor.execute("PRAGMA table_info(pivot_crm)")
                db_columns = [x[1] for x in cursor.fetchall()]
                non_element_columns = ['CRM ID', 'Solution Label', 'Analysis Method', 'Type']

                # Group by CRM ID
                crm_options = {}
                for row in crm_data:
                    crm_id = row[db_columns.index('CRM ID')]
                    crm_options[crm_id] = []
                    for col in db_columns:
                        if col in non_element_columns:
                            continue
                        value = row[db_columns.index(col)]
                        if value is not None and value != '':
                            try:
                                # Extract element symbol (e.g., 'Cu' or 'Fe')
                                symbol = col.split('_')[0].strip()
                                # If column is an oxide, map to element
                                for el, (oxide_formula, _) in oxide_factors.items():
                                    if col == oxide_formula:
                                        symbol = el
                                        break
                                crm_options[crm_id].append((symbol, float(value)))
                            except (ValueError, TypeError):
                                self.logger.warning(f"Invalid value for {col}: {value}")
                                continue

                # If multiple CRMs, show dialog for user selection
                selected_crm_id = crm_id_string
                if len(crm_options) > 1:
                    dialog = QDialog(self.pivot_tab)
                    dialog.setWindowTitle(f"Select CRM for {label}")
                    layout = QVBoxLayout(dialog)
                    layout.addWidget(QLabel(f"Multiple CRMs found for {label}. Please select one:"))
                    radio_buttons = []
                    for crm_id in crm_options.keys():
                        rb = QRadioButton(crm_id)
                        radio_buttons.append(rb)
                        layout.addWidget(rb)
                    radio_buttons[0].setChecked(True)  # Default to first option
                    confirm_btn = QPushButton("Confirm")
                    layout.addWidget(confirm_btn)

                    def on_confirm():
                        nonlocal selected_crm_id
                        for rb in radio_buttons:
                            if rb.isChecked():
                                selected_crm_id = rb.text()
                                break
                        dialog.accept()

                    confirm_btn.clicked.connect(on_confirm)
                    dialog.exec()

                print(f"Selected CRM ID: {selected_crm_id}")
                crm_data = crm_options.get(selected_crm_id, [])
                if not crm_data:
                    self.logger.warning(f"No data for selected {selected_crm_id}")
                    continue

                # Build dictionary of element symbols and grades
                crm_dict = {symbol: grade for symbol, grade in crm_data}
                print(f"CRM Dict: {crm_dict}")

                # Map CRM values to pivot_data columns using element-to-wavelength mapping
                crm_values = {'Solution Label': selected_crm_id}
                for element, columns in element_to_columns.items():
                    if element in crm_dict:
                        value = crm_dict[element]
                        if self.pivot_tab.use_oxide_var.isChecked():
                            # Apply oxide factor if column is an oxide formula
                            for col in columns:
                                for el, (oxide_formula, factor) in oxide_factors.items():
                                    if col == oxide_formula and element == el:
                                        crm_values[col] = value * factor
                                        print(f"Matched oxide {col} to {element}, value: {crm_values[col]}")
                                        break
                                else:
                                    crm_values[col] = value
                                    print(f"Matched {col} to {element}, value: {value}")
                        else:
                            # Assign value to all wavelength columns for the element
                            for col in columns:
                                crm_values[col] = value
                                print(f"Matched {col} to {element}, value: {value}")
                    else:
                        print(f"No match for element {element} in crm_dict")

                print(f"CRM Values: {crm_values}")
                # Only add if there are values beyond Solution Label
                if len(crm_values) > 1:
                    self.pivot_tab._inline_crm_rows[label] = [crm_values]
                    self.pivot_tab.included_crms[label] = QCheckBox(label, checked=True)
                else:
                    self.logger.warning(f"No matching elements for {selected_crm_id}, crm_values: {crm_values}")

            print(f"Inline CRM Rows: {self.pivot_tab._inline_crm_rows}")
            if not self.pivot_tab._inline_crm_rows:
                QMessageBox.information(self.pivot_tab, "Info", "No matching CRM elements found for comparison!")
                self.logger.info("No matching CRM elements found")
                return

            self.pivot_tab._inline_crm_rows_display = self._build_crm_row_lists_for_columns(
                list(self.pivot_tab.pivot_data.columns)
            )
            print(f"CRM Display: {self.pivot_tab._inline_crm_rows_display}")
            self.pivot_tab.update_pivot_display()

        except Exception as e:
            self.logger.error(f"Failed to check RM: {str(e)}")
            QMessageBox.warning(self.pivot_tab, "Error", f"Failed to check RM: {str(e)}")

    def _build_crm_row_lists_for_columns(self, columns):
        """Build CRM row lists for display, including differences and tags."""
        crm_display = {}
        dec = int(self.pivot_tab.decimal_places.currentText())
        try:
            min_diff = float(self.pivot_tab.diff_min.text())
            max_diff = float(self.pivot_tab.diff_max.text())
        except ValueError:
            min_diff, max_diff = -12, 12
            self.logger.warning("Invalid diff_min or diff_max, using defaults -12 and 12")

        has_act_vol = 'Act Vol' in self.pivot_tab.original_df.columns if self.pivot_tab.original_df is not None else False
        has_act_wg = 'Act Wgt' in self.pivot_tab.original_df.columns if self.pivot_tab.original_df is not None else False
        has_coeff_1 = 'Coeff 1' in self.pivot_tab.original_df.columns if self.pivot_tab.original_df is not None else False
        has_coeff_2 = 'Coeff 2' in self.pivot_tab.original_df.columns if self.pivot_tab.original_df is not None else False
        print(f"Input Inline CRM Rows: {self.pivot_tab._inline_crm_rows}")

        for sol_label, list_of_dicts in self.pivot_tab._inline_crm_rows.items():
            crm_display[sol_label] = []
            pivot_row = self.pivot_tab.pivot_data[
                self.pivot_tab.pivot_data['Solution Label'].str.strip().str.lower() == sol_label.strip().lower()
            ]
            print(f"Checking sol_label: {sol_label}, Exists in pivot_data: {not pivot_row.empty}")
            if pivot_row.empty:
                self.logger.warning(f"No pivot row found for {sol_label}")
                continue
            pivot_values = pivot_row.iloc[0].to_dict()

            element_params = {}
            if self.pivot_tab.original_df is not None:
                sample_rows = self.pivot_tab.original_df[
                    (self.pivot_tab.original_df['Solution Label'].str.strip().str.lower() == sol_label.strip().lower()) &
                    (self.pivot_tab.original_df['Type'].isin(['Sample', 'Samp']))
                ]
                for _, row in sample_rows.iterrows():
                    element = row['Element'].split('_')[0].strip()
                    element_params[element] = {
                        'Act Vol': row['Act Vol'] if has_act_vol else 1.0,
                        'Act Wgt': row['Act Wgt'] if has_act_wg else 1.0,
                        'Coeff 1': row['Coeff 1'] if has_coeff_1 else 0.0,
                        'Coeff 2': row['Coeff 2'] if has_coeff_2 else 1.0
                    }
                print(f"Element Params for {sol_label}: {element_params}")

            for d in list_of_dicts:
                crm_row_list = []
                for col in columns:
                    if col == 'Solution Label':
                        crm_row_list.append(f"{d.get('Solution Label', sol_label)} CRM")
                    else:
                        val = d.get(col, "")
                        if pd.isna(val) or val == "":
                            crm_row_list.append("")
                        else:
                            try:
                                element_symbol = col.split()[0].strip()
                                if self.pivot_tab.use_oxide_var.isChecked():
                                    for el, (oxide_formula, _) in oxide_factors.items():
                                        if col == oxide_formula:
                                            element_symbol = el
                                            break
                                params = element_params.get(element_symbol, {
                                    'Act Vol': 1.0, 'Act Wgt': 1.0, 'Coeff 1': 0.0, 'Coeff 2': 1.0
                                })
                                act_vol = params['Act Vol']
                                act_wg = params['Act Wgt']
                                coeff_1 = params['Coeff 1']
                                coeff_2 = params['Coeff 2']
                                if self.pivot_tab.use_int_var.isChecked() and act_vol != 0 and act_wg != 0:
                                    int_value = coeff_2 * (float(val) / (act_vol / act_wg)) + coeff_1
                                    if element_symbol in oxide_factors and self.pivot_tab.use_oxide_var.isChecked():
                                        _, factor = oxide_factors[element_symbol]
                                        oxide_int_value = int_value * factor
                                        crm_row_list.append(f"{oxide_int_value:.{dec}f}")
                                    else:
                                        crm_row_list.append(f"{int_value:.{dec}f}")
                                else:
                                    if element_symbol in oxide_factors and self.pivot_tab.use_oxide_var.isChecked():
                                        _, factor = oxide_factors[element_symbol]
                                        oxide_val = float(val) * factor
                                        crm_row_list.append(f"{oxide_val:.{dec}f}")
                                    else:
                                        crm_row_list.append(f"{float(val):.{dec}f}")
                            except Exception as e:
                                self.logger.warning(f"Error processing CRM value for {col}: {str(e)}")
                                crm_row_list.append(str(val))
                crm_display[sol_label].append((crm_row_list, ["crm"] * len(columns)))

                diff_row_list = []
                diff_tags = []
                for col in columns:
                    if col == 'Solution Label':
                        diff_row_list.append(f"{sol_label} Diff (%)")
                        diff_tags.append("diff")
                    else:
                        pivot_val = pivot_values.get(col, None)
                        crm_val = d.get(col, None)
                        if pivot_val is not None and crm_val is not None:
                            try:
                                element_symbol = col.split()[0].strip()
                                if self.pivot_tab.use_oxide_var.isChecked():
                                    for el, (oxide_formula, _) in oxide_factors.items():
                                        if col == oxide_formula:
                                            element_symbol = el
                                            break
                                params = element_params.get(element_symbol, {
                                    'Act Vol': 1.0, 'Act Wgt': 1.0, 'Coeff 1': 0.0, 'Coeff 2': 1.0
                                })
                                act_vol = params['Act Vol']
                                act_wg = params['Act Wgt']
                                coeff_1 = params['Coeff 1']
                                coeff_2 = params['Coeff 2']
                                pivot_val = float(pivot_val)
                                crm_val = float(crm_val)
                                if self.pivot_tab.use_int_var.isChecked() and act_vol != 0 and act_wg != 0:
                                    crm_val = coeff_2 * (crm_val / (act_vol / act_wg)) + coeff_1
                                if crm_val != 0:
                                    diff = ((crm_val - pivot_val) / crm_val) * 100
                                    diff_row_list.append(f"{diff:.{dec}f}")
                                    diff_tags.append("in_range" if min_diff <= diff <= max_diff else "out_range")
                                else:
                                    diff_row_list.append("N/A")
                                    diff_tags.append("diff")
                            except Exception as e:
                                self.logger.warning(f"Error calculating diff for {col}: {str(e)}")
                                diff_row_list.append("")
                                diff_tags.append("diff")
                        else:
                            diff_row_list.append("")
                            diff_tags.append("diff")
                crm_display[sol_label].append((diff_row_list, diff_tags))

        return crm_display