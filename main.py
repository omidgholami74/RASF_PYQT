import sys
from PyQt6.QtWidgets import QApplication
from app import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.setWindowTitle("RASF PROCESSING")
    window.setGeometry(100, 100, 900, 600)
    window.show()
    sys.exit(app.exec())