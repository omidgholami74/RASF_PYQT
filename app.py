import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFrame,QTabWidget
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
import os
import pandas as pd
import logging

# Setup logging
logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize data and file path
        self.data = None
        self.file_path = None
        self.file_path_label = QLabel("File Path: No file selected")
        
        # Initialize tabs
        self.pivot_tab = PivotTab(self, self)
        self.elements_tab = ElementsTab(self, self)
        self.crm_tab = CRMTab(self, self)
        self.results = ResultsFrame(self, self)
        self.rm_check=CheckRMFrame(self,self)
        self.weight_check =WeightCheckFrame(self,self)
        self.volume_check=VolumeCheckFrame(self,self)
        self.df_check=DFCheckFrame(self,self)
        # Tab definitions
        tab_info = {
            "File": {
                "New": "File -> New Content",
                "Open": self.handle_excel,  # Updated to use handle_excel
                "Close": "File -> Save Content"
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
                "Volume Check" : self.volume_check,
                "DF check": self.df_check,
                "RM check": self.rm_check,
                "Result": self.results
            }
        }
        
        # Create MainTabContent with tab_info
        self.main_content = MainTabContent(tab_info)
        self.setCentralWidget(self.main_content)
        self.setWindowTitle("RASF Data Processor")
    
    def handle_excel(self):
        """Load Excel/CSV file, store DataFrame, and update window title with file name"""
        result = load_excel(self)
        if result:
            self.data, self.file_path = result
            file_name = os.path.basename(self.file_path)
            self.setWindowTitle(f"RASF Data Processor - {file_name}")
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
            logger.error(f"Error in set_data: {e}")


    def notify_data_changed(self):
        """Notify all tabs that data has changed."""
        try:
            for tab_index in range(self.tabs.count()):
                tab_widget = self.tabs.widget(tab_index)
                if isinstance(tab_widget, QTabWidget):
                    for sub_tab_index in range(tab_widget.count()):
                        sub_tab_widget = tab_widget.widget(sub_tab_index)
                        if hasattr(sub_tab_widget, 'data_changed'):
                            sub_tab_widget.data_changed()
            logger.debug("Notified all tabs of data change")
        except Exception as e:
            logger.error(f"Error in notify_data_changed: {e}")
    
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