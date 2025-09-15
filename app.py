import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt
from tab import MainTabContent
from screens.calibration_tab import ElementsTab
from screens.pivot.pivot_tab import PivotTab
from screens.CRM import CRMTab
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Placeholder get_data method
        self.data = None  # Replace with actual DataFrame if needed
        
        # Tab definitions
        tab_info = {
            "File": {
                "New": "File -> New Content",
                "Open": "File -> Open Content",
                "Save": "File -> Save Content"
            },
            "Insert": {
                "Picture": PivotTab(self,self),
                "Table": CRMTab(self,self),
                "Chart": "Insert -> Chart Content"
            },
            "View": {
                "Zoom": ElementsTab(self, self),  # Pass MainWindow as app and parent
                "Layout": "View -> Layout Content"
            }
        }
        
        # Create MainTabContent with tab_info
        self.main_content = MainTabContent(tab_info)
        self.setCentralWidget(self.main_content)
    
    def get_data(self):
        """Placeholder method to return a pandas DataFrame"""
        # Replace with actual implementation
        return self.data

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())