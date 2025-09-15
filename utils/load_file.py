import pandas as pd
import csv
import os
import logging
from PyQt6.QtWidgets import QFileDialog, QMessageBox

# Setup logging
logger = logging.getLogger(__name__)

def load_excel(app):
    """Load and parse Excel/CSV file, update UI via MainTabContent, and return DataFrame and file path"""
    logger.debug("Loading Excel/CSV file")
    file_path, _ = QFileDialog.getOpenFileName(
        app,
        "Open File",
        "",
        "CSV files (*.csv);;Excel files (*.xlsx *.xls)"
    )
    
    if not file_path:
        app.file_path_label.setText("File Path: No file selected")
        app.setWindowTitle("RASF Data Processor")
        return None
        
    try:
        app.file_path_label.setText(f"File Path: {file_path}")
        
        is_new_format = False
        if file_path.endswith('.csv'):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    preview_lines = [f.readline().strip() for _ in range(10)]
                logger.debug(f"Preview lines: {preview_lines}")
                is_new_format = any("Sample ID:" in line for line in preview_lines) or \
                                any("Net Intensity" in line for line in preview_lines)
            except Exception as e:
                logger.warning(f"Preview read failed: {str(e)}. Assuming new format for CSV.")
                is_new_format = True
        else:
            try:
                preview = pd.read_excel(file_path, header=None, nrows=10)
                logger.debug(f"Excel preview:\n{preview.to_string()}")
                is_new_format = any(preview[0].str.contains("Sample ID:", na=False)) or \
                                any(preview[0].str.contains("Net Intensity", na=False))
            except Exception as e:
                logger.error(f"Failed to read Excel preview: {str(e)}")
                raise
        
        data_rows = []
        current_sample = None

        if is_new_format:
            logger.debug("Detected new file format (Sample ID-based)")
            if file_path.endswith('.csv'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        reader = list(csv.reader(f, delimiter=',', quotechar='"'))
                        total_rows = len(reader)
                        for idx, row in enumerate(reader):
                            if idx == total_rows - 1:
                                logger.debug("Skipping last row of CSV")
                                continue
                            if not row or all(cell.strip() == "" for cell in row):
                                continue
                            
                            if len(row) > 0 and row[0].startswith("Sample ID:"):
                                current_sample = row[1]
                                logger.debug(f"Found Sample ID: {current_sample}")
                                continue
                            
                            if len(row) > 0 and (row[0].startswith("Method File:") or row[0].startswith("Calibration File:")):
                                continue
                            
                            if current_sample is None:
                                current_sample = "Unknown_Sample"
                            
                            element = row[0].strip()
                            try:
                                intensity = float(row[1]) if len(row) > 1 and row[1].strip() else None
                                concentration = float(row[5]) if len(row) > 5 and row[5].strip() else None
                                if intensity is not None or concentration is not None:
                                    type_value = "Blk" if "BLANK" in current_sample.upper() else "Sample"
                                    data_rows.append({
                                        "Solution Label": current_sample,
                                        "Element": element,
                                        "Int": intensity,
                                        "Corr Con": concentration,
                                        "Type": type_value
                                    })
                                    logger.debug(f"Parsed row -> {current_sample} | {element} | {intensity} | {concentration}")
                            except Exception as e:
                                logger.warning(f"Invalid data for element {element} in sample {current_sample}: {str(e)}")
                                continue
                except Exception as e:
                    logger.error(f"Failed to parse CSV: {str(e)}")
                    raise
            else:
                try:
                    raw_data = pd.read_excel(file_path, header=None)
                    total_rows = raw_data.shape[0]
                    for index, row in raw_data.iterrows():
                        if index == total_rows - 1:
                            logger.debug("Skipping last row of Excel")
                            continue
                        row_list = row.tolist()
                        
                        if any("No valid data found in the file" in str(cell) for cell in row_list):
                            continue
                        
                        if isinstance(row[0], str) and row[0].startswith("Sample ID:"):
                            current_sample = row[0].split("Sample ID:")[1].strip()
                            logger.debug(f"Found Sample ID: {current_sample}")
                            continue
                        
                        if isinstance(row[0], str) and (row[0].startswith("Method File:") or row[0].startswith("Calibration File:")):
                            continue
                        
                        if current_sample and pd.notna(row[0]):
                            element = str(row[0]).strip()
                            try:
                                intensity = float(row[1]) if pd.notna(row[1]) else None
                                concentration = float(row[5]) if pd.notna(row[5]) else None
                                if intensity is not None or concentration is not None:
                                    type_value = "Blk" if "BLANK" in current_sample.upper() else "Sample"
                                    data_rows.append({
                                        "Solution Label": current_sample,
                                        "Element": element,
                                        "Int": intensity,
                                        "Corr Con": concentration,
                                        "Type": type_value
                                    })
                                    logger.debug(f"Parsed row -> {current_sample} | {element} | {intensity} | {concentration}")
                            except Exception as e:
                                logger.warning(f"Invalid data for element {element} in sample {current_sample}: {str(e)}")
                                continue
                except Exception as e:
                    logger.error(f"Failed to parse Excel: {str(e)}")
                    raise
        
        else:
            logger.debug("Detected previous file format (tabular)")
            if file_path.endswith('.csv'):
                try:
                    temp_df = pd.read_csv(file_path, header=None, nrows=1, on_bad_lines='skip')
                    if temp_df.iloc[0].notna().sum() == 1:
                        df = pd.read_csv(file_path, header=1, on_bad_lines='skip')
                    else:
                        df = pd.read_csv(file_path, header=0, on_bad_lines='skip')
                except Exception as e:
                    logger.error(f"Failed to read CSV as tabular: {str(e)}")
                    raise ValueError("Could not parse CSV as tabular format")
            else:
                try:
                    temp_df = pd.read_excel(file_path, header=None, nrows=1)
                    if temp_df.iloc[0].notna().sum() == 1:
                        df = pd.read_excel(file_path, header=1)
                    else:
                        df = pd.read_excel(file_path, header=0)
                except Exception as e:
                    logger.error(f"Failed to read Excel as tabular: {str(e)}")
                    raise ValueError("Could not parse Excel as tabular format")
            
            df = df.iloc[:-1]
            
            expected_columns = ["Solution Label", "Element", "Int", "Corr Con"]
            column_mapping = {"Sample ID": "Solution Label"}
            df.rename(columns=column_mapping, inplace=True)
            
            if not all(col in df.columns for col in expected_columns):
                logger.error(f"Missing columns in tabular format: {set(expected_columns) - set(df.columns)}")
                raise ValueError(f"Required columns missing: {', '.join(set(expected_columns) - set(df.columns))}")
            
            if 'Type' not in df.columns:
                df['Type'] = df['Solution Label'].apply(lambda x: "Blk" if "BLANK" in str(x).upper() else "Sample")
        
        if not data_rows and is_new_format:
            logger.error("No valid data rows were parsed")
            raise ValueError("No valid data found in the file")
        elif is_new_format:
            df = pd.DataFrame(data_rows, columns=["Solution Label", "Element", "Int", "Corr Con", "Type"])
        
        logger.debug(f"Final DataFrame shape: {df.shape}")
        
        # Store DataFrame and file path
        app.data = df
        app.file_path = file_path
        
        # Update UI via MainTabContent
        if hasattr(app, 'main_content'):
            # Update Elements tab
            if hasattr(app, 'elements_tab') and app.elements_tab:
                app.elements_tab.process_blk_elements()
            else:
                if hasattr(app, 'elements_tab'):
                    app.elements_tab.display_elements(["Cu", "Zn", "Fe"])
            
            # Trigger pivot tab updates
            if "pivot" in app.main_content.tab_subtab_map:
                pivot_subtabs = app.main_content.tab_subtab_map["pivot"]["widgets"]
                if "Display" in pivot_subtabs:
                    # Ensure pivot_tab is reset and updated
                    if hasattr(app.pivot_tab, 'reset_cache'):
                        app.main_content.switch_subtab("Display", "pivot")  # Show pivot table
                        app.pivot_tab.reset_cache()
                    if hasattr(app.pivot_tab, 'create_pivot'):
                        app.pivot_tab.create_pivot()
            
            # Trigger CRM tab updates
            if "CRM" in app.main_content.tab_subtab_map:
                crm_subtabs = app.main_content.tab_subtab_map["CRM"]["widgets"]
                if "CRM" in crm_subtabs:
                    if hasattr(app.crm_tab, 'reset_cache'):
                        app.main_content.switch_subtab("CRM", "CRM")  # Show CRM tab
                        app.crm_tab.reset_cache()
                    if hasattr(app.crm_tab, 'load_and_display'):
                        app.crm_tab.load_and_display()
            
            # Trigger Process tab updates
            if "Process" in app.main_content.tab_subtab_map:
                process_subtabs = app.main_content.tab_subtab_map["Process"]["widgets"]
                if "Result" in process_subtabs:
                    if hasattr(app.results, 'reset_cache'):
                        app.main_content.switch_subtab("Result", "Process")  # Show ResultsFrame
                        app.results.reset_cache()
                    if hasattr(app.results, 'show_processed_data'):
                        app.results.show_processed_data()
        
        # Switch to Calibration tab
        if hasattr(app, 'main_content'):
            app.main_content.switch_tab("Calibration")
        
        QMessageBox.information(app, "Success", "File loaded successfully!")
        return df, file_path
        
    except Exception as e:
        logger.error(f"Failed to load file: {str(e)}")
        QMessageBox.warning(app, "Error", f"Failed to load file:\n{str(e)}")
        app.file_path_label.setText("File Path: No file selected")
        app.setWindowTitle("RASF Data Processor")
        return None