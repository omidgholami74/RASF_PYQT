import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFrame
from PyQt6.QtCore import Qt
from tab import MainTabContent, RibbonTabButton, SubTabButton, TAB_COLORS
from screens.calibration_tab import ElementsTab
from screens.pivot.pivot_tab import PivotTab
from screens.CRM import CRMTab
from utils.load_file import load_excel
from screens.process.result import ResultsFrame
import os

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
                "Weight Check": "File -> New Content",
                "DF check": "File -> Save Content",
                "RM check": "File -> Save Content",
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