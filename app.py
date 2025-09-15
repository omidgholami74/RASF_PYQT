import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt
from tab import MainTabContent

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Tab definitions
        tab_info = {
            "File": {
                "New": "File -> New Content",
                "Open": "File -> Open Content",
                "Save": "File -> Save Content"
            },
            "Insert": {
                "Picture": "Insert -> Picture Content",
                "Table": "Insert -> Table Content",
                "Chart": "Insert -> Chart Content"
            },
            "View": {
                "Zoom": "View -> Zoom Content",
                "Layout": "View -> Layout Content"
            }
        }
        
        # Create MainTabContent with tab_info
        self.main_content = MainTabContent(tab_info)
        self.setCentralWidget(self.main_content)