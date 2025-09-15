import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office-style Tabs UI")
        self.setGeometry(100, 100, 900, 600)
        
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())