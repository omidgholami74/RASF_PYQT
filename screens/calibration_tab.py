import sys
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QTreeWidget, QTreeWidgetItem, QGridLayout
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import logging

# Setup logging
logger = logging.getLogger(__name__)

class ElementsTab(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.current_element = None
        self.filtered_elements = []  # Initialize filtered_elements
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the Elements tab UI with ribbon-style aesthetics"""
        # Main layout for elements tab content
        elements_layout = QVBoxLayout()
        elements_layout.setContentsMargins(10, 10, 10, 10)
        elements_layout.setSpacing(10)
        self.setLayout(elements_layout)
        self.setStyleSheet("background-color: white;")
        
        # Container for element buttons
        self.elements_container = QWidget(self)
        self.elements_grid_layout = QGridLayout()  # Use QGridLayout instead of QHBoxLayout
        self.elements_grid_layout.setContentsMargins(10, 10, 10, 10)
        self.elements_grid_layout.setSpacing(10)
        self.elements_container.setLayout(self.elements_grid_layout)
        self.elements_container.setStyleSheet("background-color: white;")
        elements_layout.addWidget(self.elements_container)
        
        # Filter frame for wavelength selection
        self.filter_frame = QWidget(self)
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(10, 0, 10, 10)
        filter_layout.setSpacing(5)
        filter_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.filter_frame.setLayout(filter_layout)
        self.filter_frame.setStyleSheet("background-color: white;")
        elements_layout.addWidget(self.filter_frame)
        
        # Wavelength filter dropdown
        filter_label = QLabel("Filter by Wavelength:", self.filter_frame)
        filter_label.setFont(QFont("Segoe UI", 14))
        filter_label.setStyleSheet("color: black;")
        filter_layout.addWidget(filter_label)
        
        self.wavelength_combo = QComboBox(self.filter_frame)
        self.wavelength_combo.addItem("All Wavelengths")
        self.wavelength_combo.setFont(QFont("Segoe UI", 14))
        self.wavelength_combo.setStyleSheet("""
            QComboBox {
                background-color: #dddddd;
                color: black;
                border: none;
                padding: 5px;
            }
            QComboBox:hover {
                background-color: #cccccc;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox QAbstractItemView {
                background-color: #dddddd;
                color: black;
                selection-background-color: #cccccc;
            }
        """)
        self.wavelength_combo.setFixedHeight(30)
        self.wavelength_combo.currentTextChanged.connect(self.filter_by_wavelength)
        filter_layout.addWidget(self.wavelength_combo)
        
        # Details frame with TreeWidget for element data
        self.details_frame = QWidget(self)
        details_layout = QVBoxLayout()
        details_layout.setContentsMargins(10, 10, 10, 10)
        details_layout.setSpacing(0)
        self.details_frame.setLayout(details_layout)
        self.details_frame.setStyleSheet("background-color: white;")
        elements_layout.addWidget(self.details_frame, 1)
        
        # TreeWidget for element details
        self.details_tree = QTreeWidget(self.details_frame)
        self.details_tree.setHeaderLabels(["Solution Label", "Element", "Soln Conc", "Wavelength"])
        self.details_tree.setStyleSheet("""
            QTreeWidget {
                background-color: white;
                color: black;
                font: 12px 'Segoe UI';
            }
            QTreeWidget::item {
                height: 25px;
            }
            QTreeWidget::item:selected {
                background-color: #cccccc;
                color: black;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: black;
                font: bold 12px 'Segoe UI';
                padding: 5px;
                border: none;
            }
        """)
        self.details_tree.setRootIsDecorated(False)
        self.details_tree.setSortingEnabled(True)
        self.details_tree.header().setSectionsClickable(True)
        
        # Configure columns
        columns = [
            ("Solution Label", 150),
            ("Element", 100),
            ("Soln Conc", 100),
            ("Wavelength", 100)
        ]
        
        for col, width in columns:
            col_index = columns.index((col, width))
            self.details_tree.header().resizeSection(col_index, width)
            self.details_tree.header().setSectionResizeMode(col_index, self.details_tree.header().ResizeMode.Fixed)
        
        # Add scrollbars
        self.details_tree.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.details_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        details_layout.addWidget(self.details_tree)
        
        # Ensure the frame expands
        elements_layout.setStretch(2, 1)
    
    def display_elements(self, elements):
        """Display element buttons in the container"""
        # Clear existing widgets without replacing the layout
        for i in reversed(range(self.elements_grid_layout.count())):
            item = self.elements_grid_layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()
        
        num_columns = 12
        for i, element in enumerate(elements):
            row = i // num_columns
            col = i % num_columns
            
            btn = QPushButton(element, self.elements_container)
            btn.setFixedSize(80, 30)
            btn.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #dddddd;
                    color: black;
                    border: none;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #cccccc;
                }
            """)
            btn.clicked.connect(lambda checked, el=element: self.show_element_details(el))
            self.elements_grid_layout.addWidget(btn, row, col)
            
            self.elements_grid_layout.setColumnStretch(col, 1)
            self.elements_grid_layout.setRowStretch(row, 1)
    
    def show_element_details(self, element):
        """Show details for the selected element"""
        self.current_element = element
        df = self.app.get_data()
        if df is None:
            return
            
        self.details_tree.clear()
        
        std_data = df[(df['Type'] == 'Std') & (df['Element'].str.startswith(element + ' '))]
        
        if std_data.empty:
            item = QTreeWidgetItem(["No STD data found", element, "", ""])
            self.details_tree.addTopLevelItem(item)
            self.wavelength_combo.clear()
            self.wavelength_combo.addItem("All Wavelengths")
        else:
            wavelengths = std_data['Element'].str.replace(element + ' ', '').unique()
            wavelengths = [w for w in wavelengths if w]
            
            self.wavelength_combo.clear()
            self.wavelength_combo.addItems(["All Wavelengths"] + sorted(wavelengths))
            self.wavelength_combo.setCurrentText("All Wavelengths")
            
            for _, row in std_data.iterrows():
                element_full = row.get('Element', '')
                wavelength = element_full.replace(element + ' ', '') if element_full.startswith(element + ' ') else row.get('Wavelength', '')
                item = QTreeWidgetItem([
                    str(row.get('Solution Label', '')),
                    element,
                    str(row.get('Soln Conc', '')),
                    wavelength
                ])
                self.details_tree.addTopLevelItem(item)
    
    def filter_by_wavelength(self, selected_wavelength):
        """Filter data by selected wavelength"""
        if not self.current_element or selected_wavelength == "All Wavelengths":
            self.show_element_details(self.current_element)
            return
            
        df = self.app.get_data()
        if df is None:
            return
            
        self.details_tree.clear()
        
        element_with_wavelength = f"{self.current_element} {selected_wavelength}"
        std_data = df[(df['Type'] == 'Std') & (df['Element'] == element_with_wavelength)]
        
        if std_data.empty:
            item = QTreeWidgetItem([f"No data for {selected_wavelength}", self.current_element, "", selected_wavelength])
            self.details_tree.addTopLevelItem(item)
        else:
            for _, row in std_data.iterrows():
                item = QTreeWidgetItem([
                    str(row.get('Solution Label', '')),
                    self.current_element,
                    str(row.get('Soln Conc', '')),
                    selected_wavelength
                ])
                self.details_tree.addTopLevelItem(item)
    
    def process_blk_elements(self):
        """Process BLK data and display unique elements"""
        logger.debug("Processing BLK elements")
        df = self.app.get_data()
        if df is None:
            logger.debug("No data available in process_blk_elements")
            self.display_elements(["Cu", "Zn", "Fe"])  # Fallback elements
            return
        
        # Filter BLK data
        if 'Type' in df.columns:
            blk_data = df[df['Type'] == 'Blk']
        else:
            blk_data = df[df['Solution Label'].str.contains("BLANK", case=False, na=False)]
        
        # Extract unique element names
        self.filtered_elements = []
        for element in blk_data['Element'].unique():
            element_name = str(element).split()[0]  # Get base element name (e.g., 'Cu' from 'Cu 324.754')
            if element_name and element_name not in self.filtered_elements:
                self.filtered_elements.append(element_name)
        
        # Display elements
        if self.filtered_elements:
            logger.debug(f"Displaying filtered elements: {self.filtered_elements}")
            self.display_elements(self.filtered_elements)
        else:
            logger.debug("No BLK elements found, displaying default elements")
            self.display_elements(["Cu", "Zn", "Fe"])  # Fallback if no BLK elements