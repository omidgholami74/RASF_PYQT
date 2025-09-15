import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFrame, QTabWidget
from PyQt6.QtCore import Qt
from tab import MainTabContent, RibbonTabButton, SubTabButton, TAB_COLORS
from screens.calibration_tab import ElementsTab
from screens.pivot.pivot_tab import PivotTab
from screens.CRM import CRMTab
from utils.load_file import load_excel
from screens.process.result import ResultsFrame
from screens.process.RM_check import CheckRMFrame
from screens.process.weight_check import WeightCheckFrame
from screens.process.volume_check import VolumeCheckFrame
from screens.process.DF_check import DFCheckFrame
from screens.compare_tab import CompareTab
import os
import pandas as pd
import logging

# Setup logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG)  # Set to DEBUG for more detailed logs

class MainWindow(QMainWindow):
    # Class-level list to track open windows
    open_windows = []

    def __init__(self):
        super().__init__()
        logger.debug("Creating new MainWindow instance")
        
        # Initialize data and file path
        self.data = None
        self.file_path = None
        self.file_path_label = QLabel("File Path: No file selected")
        
        # Initialize tabs only once
        self.pivot_tab = PivotTab(self, self)
        self.elements_tab = ElementsTab(self, self)
        self.crm_tab = CRMTab(self, self)
        self.results = ResultsFrame(self, self)
        self.rm_check = CheckRMFrame(self, self)
        self.weight_check = WeightCheckFrame(self, self)
        self.volume_check = VolumeCheckFrame(self, self)
        self.df_check = DFCheckFrame(self, self)
        self.compare_tab = CompareTab(self, self)
        
        # Tab definitions
        tab_info = {
            "File": {
                "Open": self.handle_excel,
                 "New": self.new_window,
                "Close": self.close_window  # Updated to use close_window
            },
            "Find similarity": {
                "display": self.compare_tab
            },
            "Elements": {
                "Display": self.elements_tab,
            },
            "pivot": {
                "Display": self.pivot_tab,
                "Export": self.pivot_tab.export_pivot,
                "Check CRM": self.pivot_tab.check_rm,
                "Clear CRM Inline": self.pivot_tab.clear_inline_crm,
                "Search": self.pivot_tab.open_row_filter_window,
                "Row Filter": self.pivot_tab.open_row_filter_window,
                "Column Filter": self.pivot_tab.open_column_filter_window
            },
            "CRM": {
                "CRM": self.crm_tab,
            },
            "Process": {
                "Weight Check": self.weight_check,
                "Volume Check": self.volume_check,
                "DF check": self.df_check,
                "RM check": self.rm_check,
                "Result": self.results
            }
        }
        
        # Create MainTabContent with tab_info
        self.main_content = MainTabContent(tab_info)
        self.setCentralWidget(self.main_content)
        self.setWindowTitle("RASF Data Processor")
        
        # Add window to open_windows list
        MainWindow.open_windows.append(self)
    
    def new_window(self):
        """Create and show a new instance of MainWindow"""
        try:
            new_window = MainWindow()
            new_window.show()
            logger.debug(f"New window created. Total open windows: {len(MainWindow.open_windows)}")
        except Exception as e:
            logger.error(f"Error creating new window: {str(e)}")

    def close_window(self):
        """Close the current window and clean up"""
        try:
            self.close()
            logger.debug(f"Closing window. Total open windows: {len(MainWindow.open_windows)}")
        except Exception as e:
            logger.error(f"Error closing window: {str(e)}")
    
    def closeEvent(self, event):
        """Handle window close event to clean up resources"""
        try:
            MainWindow.open_windows.remove(self)
            logger.debug(f"Window closed. Total open windows: {len(MainWindow.open_windows)}")
            # Ensure CRMTab closes its database connection
            if hasattr(self.crm_tab, 'close_db_connection'):
                self.crm_tab.close_db_connection()
            event.accept()
        except Exception as e:
            logger.error(f"Error in closeEvent: {str(e)}")
            event.accept()
    
    def handle_excel(self):
        """Load Excel/CSV file, store DataFrame, and update window title with file name"""
        try:
            result = load_excel(self)
            if result:
                self.data, self.file_path = result
                file_name = os.path.basename(self.file_path)
                self.setWindowTitle(f"RASF Data Processor - {file_name}")
                logger.debug(f"Excel file loaded: {file_name}")
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
    
    def set_data(self, df, for_results=False):
        """Set or update the application-wide DataFrame."""
        try:
            if not isinstance(df, pd.DataFrame):
                logger.error("Invalid data type provided to set_data: must be a pandas DataFrame")
                return
            self.data = df.copy(deep=True)
            logger.debug(f"Data set {'for results' if for_results else ''}. Shape: {self.data.shape}")
            if for_results:
                self.notify_data_changed()
        except Exception as e:
            logger.error(f"Error in set_data: {str(e)}")

    def notify_data_changed(self):
        """Notify all tabs that data has changed."""
        try:
            for tab_index in range(self.main_content.tabs.count()):
                tab_widget = self.main_content.tabs.widget(tab_index)
                if isinstance(tab_widget, QTabWidget):
                    for sub_tab_index in range(tab_widget.count()):
                        sub_tab_widget = tab_widget.widget(sub_tab_index)
                        if hasattr(sub_tab_widget, 'data_changed'):
                            sub_tab_widget.data_changed()
            logger.debug("Notified all tabs of data change")
        except Exception as e:
            logger.error(f"Error in notify_data_changed: {str(e)}")
    
    def get_data(self):
        """Return the stored DataFrame"""
        return self.data
    
    def get_excluded_samples(self):
        """Return list of excluded samples"""
        return []  # Implement as needed
    
    def get_excluded_volumes(self):
        """Return list of excluded volumes"""
        return []  # Implement as needed
    
    def get_excluded_dfs(self):
        """Return list of excluded DataFrames"""
        return []  # Implement as needed

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())