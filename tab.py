from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
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
                text-align: left;
            }}
            QPushButton:hover {{
                background-color: #dddddd;
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
                text-align: left;
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
                text-align: left;
            }}
            QPushButton:hover {{
                background-color: #dddddd;
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
                text-align: left;
            }}
            QPushButton:hover {{
                background-color: #eeeeee;
            }}
        """)
        
    def select(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: #dddddd;
                color: black;
                font: 14px 'Segoe UI';
                border: none;
                padding: 5px;
                text-align: left;
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
                text-align: left;
            }}
            QPushButton:hover {{
                background-color: #eeeeee;
            }}
        """)

class MainTabContent(QWidget):
    def __init__(self, tab_info, parent=None):
        super().__init__(parent)
        self.current_tab = None
        self.tab_subtab_map = {}

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Ribbon bar
        self.ribbon = QWidget()
        self.ribbon.setFixedHeight(30)
        self.ribbon.setStyleSheet("background-color: white;")
        ribbon_layout = QHBoxLayout()
        ribbon_layout.setContentsMargins(0, 0, 0, 0)
        ribbon_layout.setSpacing(0)
        ribbon_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.ribbon.setLayout(ribbon_layout)
        
        # Content frame
        self.content_frame = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        self.content_frame.setLayout(content_layout)
        
        main_layout.addWidget(self.ribbon)
        main_layout.addWidget(self.content_frame, 1)
        
        # Tab definitions
        self.tabs = {}
        self.tab_buttons = {}
        
        # Create tabs and buttons
        for t_name, subtabs in tab_info.items():
            colors = TAB_COLORS.get(t_name, {"bg": "white", "indicator": "black"})
            btn = RibbonTabButton(t_name, bg_color=colors["bg"], text_color="black")
            btn.clicked.connect(lambda checked, n=t_name: self.switch_tab(n))
            ribbon_layout.addWidget(btn)
            self.tab_buttons[t_name] = btn
            
            # Create tab content
            tab_content = QWidget()
            tab_content_layout = QVBoxLayout()
            tab_content_layout.setContentsMargins(0, 0, 0, 0)
            tab_content_layout.setSpacing(0)
            tab_content.setLayout(tab_content_layout)
            self.tabs[t_name] = tab_content
            content_layout.addWidget(tab_content)
            tab_content.hide()
            
            # Subtab bar
            subtab_bar = QWidget()
            subtab_bar.setFixedHeight(50)
            subtab_bar.setStyleSheet(f"background-color: {colors['bg']};")
            subtab_layout = QHBoxLayout()
            subtab_layout.setContentsMargins(8, 6, 8, 6)
            subtab_layout.setSpacing(8)
            subtab_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
            subtab_bar.setLayout(subtab_layout)
            
            # Indicator
            indicator = QWidget()
            indicator.setFixedHeight(3)
            indicator.setStyleSheet(f"background-color: {colors['indicator']};")
            
            # Subtab content area
            subtab_content = QWidget()
            subtab_content_layout = QVBoxLayout()
            subtab_content_layout.setContentsMargins(0, 0, 0, 0)
            subtab_content_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
            subtab_content.setLayout(subtab_content_layout)
            
            tab_content_layout.addWidget(subtab_bar)
            tab_content_layout.addWidget(indicator)
            tab_content_layout.addWidget(subtab_content, 1)
            
            # Subtab buttons and content
            subtab_buttons = {}
            subtab_widgets = {}
            
            for name, content in subtabs.items():
                btn = SubTabButton(name, text_color="black")
                btn.clicked.connect(lambda checked, n=name, tn=t_name: self.switch_subtab(n, tn))
                subtab_layout.addWidget(btn)
                subtab_buttons[name] = btn
                
                # Create subtab content
                if isinstance(content, str):
                    frame = QWidget()
                    frame_layout = QVBoxLayout()
                    frame_layout.setContentsMargins(25, 25, 25, 25)
                    frame_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
                    label = QLabel(content)
                    label.setStyleSheet("font: 18px 'Segoe UI'; color: black;")
                    frame_layout.addWidget(label)
                    frame.setLayout(frame_layout)
                else:
                    frame = content  # Use the provided widget (e.g., ElementsTab)
                subtab_widgets[name] = frame
                subtab_content_layout.addWidget(frame)
                frame.hide()
                
            self.tab_subtab_map[t_name] = {
                "buttons": subtab_buttons,
                "widgets": subtab_widgets,
                "content": subtab_content,
                "current_subtab": None
            }
            
            if subtabs:
                self.switch_subtab(list(subtabs.keys())[0], t_name)
        
        self.setLayout(main_layout)
        
        if tab_info:
            self.switch_tab(list(tab_info.keys())[0])
            
    def switch_subtab(self, name, tab_name):
        if tab_name not in self.tab_subtab_map:
            return
        
        subtab_buttons = self.tab_subtab_map[tab_name]["buttons"]
        subtab_widgets = self.tab_subtab_map[tab_name]["widgets"]
        current_subtab = self.tab_subtab_map[tab_name]["current_subtab"]
        
        if current_subtab:
            subtab_widgets[current_subtab].hide()
            subtab_buttons[current_subtab].deselect()
            
        subtab_widgets[name].show()
        subtab_buttons[name].select()
        self.tab_subtab_map[tab_name]["current_subtab"] = name
        
    def switch_tab(self, tab_name):
        if self.current_tab and self.current_tab in self.tabs:
            self.tabs[self.current_tab].hide()
            self.tab_buttons[self.current_tab].deselect()
            
        self.tabs[tab_name].show()
        self.tab_buttons[tab_name].select()
        self.current_tab = tab_name
        
        colors = TAB_COLORS.get(tab_name, {"bg": "white", "indicator": "black"})
        self.tabs[tab_name].layout().itemAt(0).widget().setStyleSheet(f"background-color: {colors['bg']};")
        self.tabs[tab_name].layout().itemAt(1).widget().setStyleSheet(f"background-color: {colors['indicator']};")