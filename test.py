import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt

# Color definitions for tabs
TAB_COLORS = {
    "File": {"bg": "#fdf6e3", "indicator": "#b58900"},
    "Insert": {"bg": "#eee8f4", "indicator": "#6a5acd"},
    "View": {"bg": "#e8f6f3", "indicator": "#2e8b57"},
}

class RibbonTabButton(QPushButton):
    def __init__(self, text, parent=None, bg_color="#fdf6e3", text_color="black"):
        super().__init__(text, parent)
        self.bg_color = bg_color
        self.text_color = text_color
        self.setFixedSize(90, 30)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.text_color};
                font: bold 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
            QPushButton:hover {{
                background-color: #ddd;
            }}
        """)
        
    def select(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.bg_color};
                color: {self.text_color};
                font: bold 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
        """)
        
    def deselect(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.text_color};
                font: bold 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
            QPushButton:hover {{
                background-color: #ddd;
            }}
        """)

class SubTabButton(QPushButton):
    def __init__(self, text, parent=None, text_color="black"):
        super().__init__(text, parent)
        self.text_color = text_color
        self.setFixedSize(110, 40)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.text_color};
                font: 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
            QPushButton:hover {{
                background-color: #eee;
            }}
        """)
        
    def select(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: #ddd;
                color: black;
                font: 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
        """)
        
    def deselect(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.text_color};
                font: 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: center;
            }}
            QPushButton:hover {{
                background-color: #eee;
            }}
        """)

class MainTabContent(QWidget):
    def __init__(self, subtabs_info, tab_name, parent=None):
        super().__init__(parent)
        self.tab_name = tab_name
        self.current_subtab = None
        
        # Main layout
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Subtab bar
        self.subtab_bar = QWidget()
        self.subtab_bar.setFixedHeight(50)
        self.subtab_bar.setStyleSheet("background-color: white;")
        subtab_layout = QHBoxLayout()
        subtab_layout.setContentsMargins(8, 6, 8, 6)
        subtab_layout.setSpacing(8)
        subtab_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)  # Left-align subtabs
        self.subtab_bar.setLayout(subtab_layout)
        
        # Indicator
        self.indicator = QWidget()
        self.indicator.setFixedHeight(3)
        self.indicator.setStyleSheet("background-color: black;")
        
        # Subtab content area
        self.subtab_content = QWidget()
        self.subtab_content_layout = QVBoxLayout()
        self.subtab_content_layout.setContentsMargins(0, 0, 0, 0)
        self.subtab_content_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)  # Left-align content
        self.subtab_content.setLayout(self.subtab_content_layout)
        
        # Add widgets to main layout
        layout.addWidget(self.subtab_bar)
        layout.addWidget(self.indicator)
        layout.addWidget(self.subtab_content, 1)
        self.setLayout(layout)
        
        # Subtab buttons and content
        self.subtab_buttons = {}
        self.subtab_widgets = {}
        
        for name, content_text in subtabs_info.items():
            # Create subtab button
            btn = SubTabButton(name, text_color="black")
            btn.clicked.connect(lambda checked, n=name: self.switch_subtab(n))
            subtab_layout.addWidget(btn)
            self.subtab_buttons[name] = btn
            
            # Create subtab content
            frame = QWidget()
            frame_layout = QVBoxLayout()
            frame_layout.setContentsMargins(25, 25, 25, 25)
            frame_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)  # Left-align content
            label = QLabel(content_text)
            label.setStyleSheet("font: 18px 'Segoe UI'; color: black;")
            frame_layout.addWidget(label)
            frame.setLayout(frame_layout)
            self.subtab_widgets[name] = frame
            
        # Show first subtab
        if subtabs_info:
            self.switch_subtab(list(subtabs_info.keys())[0])
            
    def switch_subtab(self, name):
        if self.current_subtab:
            self.subtab_widgets[self.current_subtab].hide()
            self.subtab_buttons[self.current_subtab].deselect()
            
        self.subtab_widgets[name].show()
        self.subtab_content_layout.addWidget(self.subtab_widgets[name])
        self.subtab_buttons[name].select()
        self.current_subtab = name
        
    def update_colors(self, bg_color, indicator_color):
        self.subtab_bar.setStyleSheet(f"background-color: {bg_color};")
        self.indicator.setStyleSheet(f"background-color: {indicator_color};")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office-style Tabs UI")
        self.setGeometry(100, 100, 900, 600)
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_widget.setLayout(main_layout)
        
        # Ribbon bar
        self.ribbon = QWidget()
        self.ribbon.setFixedHeight(30)
        self.ribbon.setStyleSheet("background-color: white;")
        ribbon_layout = QHBoxLayout()
        ribbon_layout.setContentsMargins(0, 0, 0, 0)
        ribbon_layout.setSpacing(0)
        ribbon_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)  # Left-align tabs
        self.ribbon.setLayout(ribbon_layout)
        
        # Content frame
        self.content_frame = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_frame.setLayout(content_layout)
        
        main_layout.addWidget(self.ribbon)
        main_layout.addWidget(self.content_frame, 1)
        
        # Tab definitions
        self.tabs = {}
        self.tab_buttons = {}
        self.current_tab = None
        
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
        
        # Create tabs and buttons
        for tab_name, subtabs in tab_info.items():
            colors = TAB_COLORS.get(tab_name, {"bg": "white", "indicator": "black"})
            btn = RibbonTabButton(tab_name, bg_color=colors["bg"], text_color="black")
            btn.clicked.connect(lambda checked, n=tab_name: self.switch_tab(n))
            ribbon_layout.addWidget(btn)
            self.tab_buttons[tab_name] = btn
            
            tab_content = MainTabContent(subtabs, tab_name)
            self.tabs[tab_name] = tab_content
            
        # Show first tab
        self.switch_tab("File")
        
    def switch_tab(self, tab_name):
        if self.current_tab:
            self.tabs[self.current_tab].hide()
            self.tab_buttons[self.current_tab].deselect()
            
        self.tabs[tab_name].show()
        self.content_frame.layout().addWidget(self.tabs[tab_name])
        self.tab_buttons[tab_name].select()
        self.current_tab = tab_name
        
        colors = TAB_COLORS.get(tab_name, {"bg": "white", "indicator": "black"})
        self.tabs[tab_name].update_colors(bg_color=colors["bg"], indicator_color=colors["indicator"])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())