import logging
import pandas as pd
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QTreeWidget, QTreeWidgetItem, QGridLayout
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush

# Setup logging (minimal for performance)
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

class ElementsTab(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.current_element = None
        self.filtered_elements = []
        self.df_cache = None  # Cache for DataFrame
        self.setup_ui()

    def setup_ui(self):
        """Setup the Elements tab UI with ribbon-style aesthetics"""
        elements_layout = QVBoxLayout()
        elements_layout.setContentsMargins(10, 10, 10, 10)
        elements_layout.setSpacing(10)
        self.setLayout(elements_layout)
        self.setStyleSheet("background-color: white;")

        # Container for element buttons
        self.elements_container = QWidget(self)
        self.elements_grid_layout = QGridLayout()
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
                background-color: #e0e0e0;
                color: black;
                border: 1px solid #aaaaaa;
                border-radius: 4px;
                padding: 5px;
            }
            QComboBox:hover {
                background-color: #d0d0d0;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                background-color: #e0e0e0;
                color: black;
                selection-background-color: #b0b0b0;
                border: 1px solid #aaaaaa;
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
                border: 1px solid #cccccc;
                border-radius: 4px;
                gridline-color: #dddddd;
            }
            QTreeWidget::item {
                height: 28px;
                padding: 2px;
                border-bottom: 1px solid #dddddd;
                border-right: 1px solid #dddddd;
            }
            QTreeWidget::item:selected {
                background-color: #b0c4de;
                color: black;
            }
            QTreeWidget::item:hover {
                background-color: #f0f0f0;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                color: black;
                font: bold 12px 'Segoe UI';
                padding: 6px;
                border: none;
                border-bottom: 1px solid #cccccc;
                border-right: 1px solid #cccccc;
            }
            QHeaderView::section:hover {
                background-color: #d0d0d0;
            }
        """)
        self.details_tree.setRootIsDecorated(False)
        self.details_tree.setSortingEnabled(True)
        self.details_tree.setAlternatingRowColors(True)  # Enable zebra striping
        self.details_tree.header().setSectionsClickable(True)

        # Configure columns
        columns = [
            ("Solution Label", 200),
            ("Element", 120),
            ("Soln Conc", 120),
            ("Wavelength", 120)
        ]
        for col, width in columns:
            col_index = columns.index((col, width))
            self.details_tree.header().resizeSection(col_index, width)
            self.details_tree.header().setSectionResizeMode(col_index, self.details_tree.header().ResizeMode.Interactive)

        # Add scrollbars
        self.details_tree.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.details_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        details_layout.addWidget(self.details_tree)
        elements_layout.setStretch(2, 1)

    def display_elements(self, elements):
        """Display element buttons in the container"""
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
                    background-color: #e0e0e0;
                    color: black;
                    border: 1px solid #aaaaaa;
                    border-radius: 4px;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #d0d0d0;
                }
                QPushButton:pressed {
                    background-color: #b0b0b0;
                }
            """)
            btn.clicked.connect(lambda checked, el=element: self.show_element_details(el))
            self.elements_grid_layout.addWidget(btn, row, col)

    def clean_dataframe(self, df):
        """Clean DataFrame to ensure valid data"""
        if df is None or not isinstance(df, pd.DataFrame):
            logger.error("Invalid DataFrame provided")
            return None
        try:
            # Select only required columns
            df = df[['Type', 'Element', 'Solution Label', 'Soln Conc']].copy()
            # Remove NaN values and convert to string
            df = df.dropna(subset=['Type', 'Element'])
            df['Type'] = df['Type'].astype(str).str.strip()
            df['Element'] = df['Element'].astype(str).str.strip()
            return df
        except Exception as e:
            logger.error(f"Error cleaning DataFrame: {str(e)}")
            return None

    def show_element_details(self, element):
        """Show details for the selected element"""
        self.current_element = element
        if self.df_cache is None:
            self.df_cache = self.clean_dataframe(self.app.get_data())
        df = self.df_cache

        if df is None:
            logger.error("No valid DataFrame available")
            self.details_tree.clear()
            item = QTreeWidgetItem(["No data available", element, "", ""])
            item.setForeground(0, QBrush(QColor("red")))
            self.details_tree.addTopLevelItem(item)
            return

        # Clear previous details
        self.details_tree.clear()

        # Get STD data for the element
        try:
            std_data = df[(df['Type'] == 'Std') & (df['Element'].str.startswith(element + ' '))]
        except Exception as e:
            logger.error(f"Error filtering STD data: {str(e)}")
            item = QTreeWidgetItem([f"Error: {str(e)}", element, "", ""])
            item.setForeground(0, QBrush(QColor("red")))
            self.details_tree.addTopLevelItem(item)
            return

        if std_data.empty:
            item = QTreeWidgetItem(["No STD data found", element, "", ""])
            item.setForeground(0, QBrush(QColor("gray")))
            self.details_tree.addTopLevelItem(item)
            self.wavelength_combo.blockSignals(True)  # Prevent recursion
            self.wavelength_combo.clear()
            self.wavelength_combo.addItem("All Wavelengths")
            self.wavelength_combo.blockSignals(False)
        else:
            # Extract unique wavelengths
            try:
                wavelengths = std_data['Element'].str.replace(element + ' ', '', regex=False).unique()
                wavelengths = [w for w in wavelengths if w]
            except Exception as e:
                logger.error(f"Error extracting wavelengths: {str(e)}")
                wavelengths = []

            # Update dropdown menu
            self.wavelength_combo.blockSignals(True)  # Prevent recursion
            self.wavelength_combo.clear()
            self.wavelength_combo.addItems(["All Wavelengths"] + sorted(wavelengths))
            self.wavelength_combo.setCurrentText("All Wavelengths")
            self.wavelength_combo.blockSignals(False)

            # Show all data
            try:
                items = [
                    QTreeWidgetItem([
                        str(row.get('Solution Label', '')),
                        element,
                        str(row.get('Soln Conc', '')),
                        row.get('Element', '').replace(element + ' ', '') if row.get('Element', '').startswith(element + ' ') else row.get('Wavelength', '')
                    ]) for _, row in std_data.iterrows()
                ]
                for i, item in enumerate(items):
                    if i % 2 == 0:  # Enhance alternating colors
                        item.setBackground(0, QBrush(QColor("#f8f8f8")))
                        item.setBackground(1, QBrush(QColor("#f8f8f8")))
                        item.setBackground(2, QBrush(QColor("#f8f8f8")))
                        item.setBackground(3, QBrush(QColor("#f8f8f8")))
                self.details_tree.addTopLevelItems(items)
            except Exception as e:
                logger.error(f"Error populating tree: {str(e)}")
                item = QTreeWidgetItem([f"Error: {str(e)}", element, "", ""])
                item.setForeground(0, QBrush(QColor("red")))
                self.details_tree.addTopLevelItem(item)

    def filter_by_wavelength(self, selected_wavelength):
        """Filter data by selected wavelength"""
        if not self.current_element or selected_wavelength == "All Wavelengths":
            self.show_element_details(self.current_element)
            return

        if self.df_cache is None:
            self.df_cache = self.clean_dataframe(self.app.get_data())
        df = self.df_cache

        if df is None:
            logger.error("No valid DataFrame available")
            self.details_tree.clear()
            item = QTreeWidgetItem(["No data available", self.current_element, "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("red")))
            self.details_tree.addTopLevelItem(item)
            return

        # Clear previous details
        self.details_tree.clear()

        # Get STD data for the element and specific wavelength
        element_with_wavelength = f"{self.current_element} {selected_wavelength}".strip()
        try:
            # Step-by-step filtering to avoid recursion error
            type_mask = df['Type'] == 'Std'
            element_mask = df['Element'] == element_with_wavelength
            std_data = df[type_mask & element_mask]
        except Exception as e:
            logger.error(f"Error filtering data: {str(e)}")
            item = QTreeWidgetItem([f"Error: {str(e)}", self.current_element, "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("red")))
            self.details_tree.addTopLevelItem(item)
            return

        if std_data.empty:
            item = QTreeWidgetItem([f"No data for {selected_wavelength}", self.current_element, "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("gray")))
            self.details_tree.addTopLevelItem(item)
        else:
            try:
                items = [
                    QTreeWidgetItem([
                        str(row.get('Solution Label', '')),
                        self.current_element,
                        str(row.get('Soln Conc', '')),
                        selected_wavelength
                    ]) for _, row in std_data.iterrows()
                ]
                for i, item in enumerate(items):
                    if i % 2 == 0:
                        item.setBackground(0, QBrush(QColor("#f8f8f8")))
                        item.setBackground(1, QBrush(QColor("#f8f8f8")))
                        item.setBackground(2, QBrush(QColor("#f8f8f8")))
                        item.setBackground(3, QBrush(QColor("#f8f8f8")))
                self.details_tree.addTopLevelItems(items)
            except Exception as e:
                logger.error(f"Error populating tree: {str(e)}")
                item = QTreeWidgetItem([f"Error: {str(e)}", self.current_element, "", selected_wavelength])
                item.setForeground(0, QBrush(QColor("red")))
                self.details_tree.addTopLevelItem(item)

    def process_blk_elements(self):
        """Process BLK data and display unique elements"""
        logger.info("Processing BLK elements")
        self.df_cache = self.clean_dataframe(self.app.get_data())
        df = self.df_cache

        if df is None:
            logger.info("No valid data available in process_blk_elements")
            self.display_elements(["Cu", "Zn", "Fe"])
            return

        # Filter BLK data
        try:
            if 'Type' in df.columns:
                blk_data = df[df['Type'] == 'Blk']
            else:
                blk_data = df[df['Solution Label'].str.contains("BLANK", case=False, na=False)]
        except Exception as e:
            logger.error(f"Error filtering BLK data: {str(e)}")
            self.display_elements(["Cu", "Zn", "Fe"])
            return

        # Extract unique element names
        self.filtered_elements = []
        try:
            for element in blk_data['Element'].unique():
                element_name = element.split()[0]
                if element_name and element_name not in self.filtered_elements:
                    self.filtered_elements.append(element_name)
        except Exception as e:
            logger.error(f"Error extracting elements: {str(e)}")
            self.display_elements(["Cu", "Zn", "Fe"])
            return

        # Display elements
        if self.filtered_elements:
            logger.info(f"Displaying filtered elements: {self.filtered_elements}")
            self.display_elements(self.filtered_elements)
        else:
            logger.info("No BLK elements found, displaying default elements")
            self.display_elements(["Cu", "Zn", "Fe"])