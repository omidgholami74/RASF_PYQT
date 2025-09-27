import logging
import pandas as pd
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel, QTreeWidget, QTreeWidgetItem, QGridLayout
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush

# Setup logging with minimal output for performance
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CustomTreeWidgetItem(QTreeWidgetItem):
    """Custom QTreeWidgetItem to enforce numeric sorting for Soln Conc and Int columns."""
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        if column in [2, 3]:  # Soln Conc or Int columns
            value_self = self.data(column, Qt.ItemDataRole.UserRole)
            value_other = other.data(column, Qt.ItemDataRole.UserRole)
            try:
                return float(value_self) < float(value_other)
            except (ValueError, TypeError):
                return str(value_self) < str(value_other)
        else:
            # For non-numeric columns, use text-based sorting
            return self.text(column).lower() < other.text(column).lower()

class ElementsTab(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app
        self.current_element = None
        self.filtered_elements = []
        self.df_cache = None  # Cache for DataFrame
        self.setup_ui()

    def setup_ui(self):
        """Setup the Elements tab UI with modern, ribbon-style aesthetics"""
        logger.info("Setting up ElementsTab UI")
        # Main layout
        elements_layout = QVBoxLayout()
        elements_layout.setContentsMargins(15, 15, 15, 15)
        elements_layout.setSpacing(12)
        self.setLayout(elements_layout)
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f6f5;
                font-family: 'Segoe UI';
            }
        """)

        # Container for element buttons
        self.elements_container = QWidget(self)
        self.elements_grid_layout = QGridLayout()
        self.elements_grid_layout.setContentsMargins(10, 10, 10, 10)
        self.elements_grid_layout.setSpacing(8)
        self.elements_container.setLayout(self.elements_grid_layout)
        self.elements_container.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
            }
        """)
        elements_layout.addWidget(self.elements_container)

        # Filter frame for wavelength selection
        self.filter_frame = QWidget(self)
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(10, 8, 10, 8)
        filter_layout.setSpacing(10)
        filter_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.filter_frame.setLayout(filter_layout)
        self.filter_frame.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 8px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
            }
        """)
        elements_layout.addWidget(self.filter_frame)

        # Wavelength filter label
        filter_label = QLabel("Filter by Wavelength:", self.filter_frame)
        filter_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        filter_label.setStyleSheet("color: #333333;")
        filter_layout.addWidget(filter_label)

        # Wavelength filter dropdown
        self.wavelength_combo = QComboBox(self.filter_frame)
        self.wavelength_combo.addItem("All Wavelengths")
        self.wavelength_combo.setFont(QFont("Segoe UI", 12))
        self.wavelength_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 6px;
                min-width: 150px;
                transition: all 0.2s ease;
            }
            QComboBox:hover {
                background-color: #f0f0f0;
                border-color: #aaaaaa;
            }
            QComboBox::drop-down {
                border: none;
                width: 24px;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #e6f3fa;
                selection-color: #333333;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }
        """)
        self.wavelength_combo.setFixedHeight(34)
        self.wavelength_combo.setToolTip("Select a wavelength to filter element data")
        self.wavelength_combo.currentTextChanged.connect(self.filter_by_wavelength)
        filter_layout.addWidget(self.wavelength_combo)
        filter_layout.addStretch()

        # Details frame with TreeWidget for element data
        self.details_frame = QWidget(self)
        details_layout = QVBoxLayout()
        details_layout.setContentsMargins(10, 10, 10, 10)
        details_layout.setSpacing(8)
        self.details_frame.setLayout(details_layout)
        self.details_frame.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
            }
        """)
        elements_layout.addWidget(self.details_frame, 1)

        # TreeWidget for element details
        self.details_tree = QTreeWidget(self.details_frame)
        self.details_tree.setHeaderLabels(["Solution Label", "Element", "Soln Conc", "Int", "Wavelength"])
        self.details_tree.setStyleSheet("""
            QTreeWidget {
                background-color: #ffffff;
                color: #333333;
                font: 12px 'Segoe UI';
                border: none;
                border-radius: 4px;
                gridline-color: #e0e0e0;
            }
            QTreeWidget::item {
                height: 30px;
                padding: 4px;
                border-bottom: 1px solid #e8e8e8;
            }
            QTreeWidget::item:selected {
                background-color: #e6f3fa;
                color: #333333;
            }
            QTreeWidget::item:hover {
                background-color: #f5f6f5;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #333333;
                font: bold 12px 'Segoe UI';
                padding: 8px;
                border: none;
                border-bottom: 1px solid #e0e0e0;
                border-right: 1px solid #e0e0e0;
            }
            QHeaderView::section:hover {
                background-color: #e8e8e8;
            }
        """)
        self.details_tree.setRootIsDecorated(False)
        self.details_tree.setSortingEnabled(True)
        self.details_tree.setAlternatingRowColors(True)
        self.details_tree.header().setSectionsClickable(True)

        # Configure columns
        columns = [
            ("Solution Label", 220),
            ("Element", 140),
            ("Soln Conc", 140),
            ("Int", 140),
            ("Wavelength", 140)
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

        # Connect sort signal
        self.details_tree.header().sortIndicatorChanged.connect(self.handle_sort)
        logger.info("UI setup completed")

    def handle_sort(self, logical_index, order):
        """Handle sorting of the QTreeWidget with proper numeric sorting for Soln Conc and Int."""
        logger.info(f"Sorting column {logical_index} with order {order}")
        self.details_tree.sortItems(logical_index, order)

    def display_elements(self, elements):
        """Display element buttons in the container"""
        logger.info(f"Displaying elements: {elements}")
        for i in reversed(range(self.elements_grid_layout.count())):
            item = self.elements_grid_layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()

        num_columns = 12
        for i, element in enumerate(elements):
            row = i // num_columns
            col = i % num_columns
            btn = QPushButton(element, self.elements_container)
            btn.setFixedSize(85, 34)
            btn.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #333333;
                    border: 1px solid #cccccc;
                    border-radius: 4px;
                    padding: 6px;
                    transition: all 0.2s ease;
                }
                QPushButton:hover {
                    background-color: #e6f3fa;
                    border-color: #4a90e2;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                }
                QPushButton:pressed {
                    background-color: #d1e7f5;
                    border-color: #2e6da4;
                }
            """)
            btn.setToolTip(f"View details for {element}")
            btn.clicked.connect(lambda checked, el=element: self.show_element_details(el))
            self.elements_grid_layout.addWidget(btn, row, col)

    def clean_dataframe(self, df):
        """Clean DataFrame to ensure valid data"""
        logger.info("Cleaning DataFrame")
        if df is None or not isinstance(df, pd.DataFrame):
            logger.error("Invalid DataFrame provided")
            return None
        try:
            df = df[['Type', 'Element', 'Solution Label', 'Soln Conc', 'Int']].copy()
            df = df.dropna(subset=['Type', 'Element'])
            df['Type'] = df['Type'].astype(str).str.strip()
            df['Element'] = df['Element'].astype(str).str.strip()
            df['Solution Label'] = df['Solution Label'].astype(str).str.strip()
            # Convert Soln Conc and Int to numeric, preserving -1.0 for invalid values
            df['Soln Conc'] = pd.to_numeric(df['Soln Conc'], errors='coerce').fillna(-1.0)
            df['Int'] = pd.to_numeric(df['Int'], errors='coerce').fillna(-1.0)
            logger.info("DataFrame cleaned successfully")
            return df
        except Exception as e:
            logger.error(f"Error cleaning DataFrame: {str(e)}")
            return None

    def show_element_details(self, element):
        """Show details for the selected element"""
        logger.info(f"Showing details for element: {element}")
        self.current_element = element
        if self.df_cache is None:
            self.df_cache = self.clean_dataframe(self.app.get_data())
        df = self.df_cache

        if df is None:
            logger.error("No valid DataFrame available")
            self.details_tree.clear()
            item = CustomTreeWidgetItem(["No data available", element, "", "", ""])
            item.setForeground(0, QBrush(QColor("#d32f2f")))
            self.details_tree.addTopLevelItem(item)
            return

        self.details_tree.clear()

        try:
            std_data = df[(df['Type'] == 'Std') & (df['Element'].str.startswith(element + ' '))]
        except Exception as e:
            logger.error(f"Error filtering STD data: {str(e)}")
            item = CustomTreeWidgetItem([f"Error: {str(e)}", element, "", "", ""])
            item.setForeground(0, QBrush(QColor("#d32f2f")))
            self.details_tree.addTopLevelItem(item)
            return

        if std_data.empty:
            logger.warning(f"No STD data found for element: {element}")
            item = CustomTreeWidgetItem(["No STD data found", element, "", "", ""])
            item.setForeground(0, QBrush(QColor("#757575")))
            self.details_tree.addTopLevelItem(item)
            self.wavelength_combo.blockSignals(True)
            self.wavelength_combo.clear()
            self.wavelength_combo.addItem("All Wavelengths")
            self.wavelength_combo.blockSignals(False)
        else:
            try:
                wavelengths = std_data['Element'].str.replace(element + ' ', '', regex=False).unique()
                wavelengths = [w for w in wavelengths if w]
            except Exception as e:
                logger.error(f"Error extracting wavelengths: {str(e)}")
                wavelengths = []

            self.wavelength_combo.blockSignals(True)
            self.wavelength_combo.clear()
            self.wavelength_combo.addItems(["All Wavelengths"] + sorted(wavelengths))
            self.wavelength_combo.setCurrentText("All Wavelengths")
            self.wavelength_combo.blockSignals(False)

            try:
                items = []
                for _, row in std_data.iterrows():
                    soln_conc = row['Soln Conc']
                    int_val = row['Int']
                    # Format display values
                    soln_conc_display = '---' if soln_conc == -1.0 else f"{soln_conc:.2f}"
                    int_display = '---' if int_val == -1.0 else f"{int_val:.2f}"
                    wavelength = row.get('Element', '').replace(element + ' ', '') if row.get('Element', '').startswith(element + ' ') else row.get('Wavelength', '')
                    item = CustomTreeWidgetItem([
                        str(row.get('Solution Label', '')),
                        element,
                        soln_conc_display,
                        int_display,
                        wavelength
                    ])
                    # Store numeric values for sorting
                    item.setData(2, Qt.ItemDataRole.UserRole, float(soln_conc))
                    item.setData(3, Qt.ItemDataRole.UserRole, float(int_val))
                    # Store text for non-numeric columns
                    item.setData(0, Qt.ItemDataRole.UserRole, str(row.get('Solution Label', '')).lower())
                    item.setData(1, Qt.ItemDataRole.UserRole, element.lower())
                    item.setData(4, Qt.ItemDataRole.UserRole, wavelength.lower())
                    items.append(item)

                for i, item in enumerate(items):
                    # Apply alternating background
                    if i % 2 == 0:
                        for col in range(5):
                            item.setBackground(col, QBrush(QColor("#fafafa")))
                self.details_tree.addTopLevelItems(items)
            except Exception as e:
                logger.error(f"Error populating tree: {str(e)}")
                item = CustomTreeWidgetItem([f"Error: {str(e)}", element, "", "", ""])
                item.setForeground(0, QBrush(QColor("#d32f2f")))
                self.details_tree.addTopLevelItem(item)

    def filter_by_wavelength(self, selected_wavelength):
        """Filter data by selected wavelength"""
        logger.info(f"Filtering by wavelength: {selected_wavelength}")
        if not self.current_element or selected_wavelength == "All Wavelengths":
            self.show_element_details(self.current_element)
            return

        if self.df_cache is None:
            self.df_cache = self.clean_dataframe(self.app.get_data())
        df = self.df_cache

        if df is None:
            logger.error("No valid DataFrame available")
            self.details_tree.clear()
            item = CustomTreeWidgetItem(["No data available", self.current_element, "", "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("#d32f2f")))
            self.details_tree.addTopLevelItem(item)
            return

        self.details_tree.clear()

        element_with_wavelength = f"{self.current_element} {selected_wavelength}".strip()
        try:
            type_mask = df['Type'] == 'Std'
            element_mask = df['Element'] == element_with_wavelength
            std_data = df[type_mask & element_mask]
        except Exception as e:
            logger.error(f"Error filtering data: {str(e)}")
            item = CustomTreeWidgetItem([f"Error: {str(e)}", self.current_element, "", "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("#d32f2f")))
            self.details_tree.addTopLevelItem(item)
            return

        if std_data.empty:
            logger.warning(f"No data found for wavelength: {selected_wavelength}")
            item = CustomTreeWidgetItem([f"No data for {selected_wavelength}", self.current_element, "", "", selected_wavelength])
            item.setForeground(0, QBrush(QColor("#757575")))
            self.details_tree.addTopLevelItem(item)
        else:
            try:
                items = []
                for _, row in std_data.iterrows():
                    soln_conc = row['Soln Conc']
                    int_val = row['Int']
                    # Format display values
                    soln_conc_display = '---' if soln_conc == -1.0 else f"{soln_conc:.2f}"
                    int_display = '---' if int_val == -1.0 else f"{int_val:.2f}"
                    item = CustomTreeWidgetItem([
                        str(row.get('Solution Label', '')),
                        self.current_element,
                        soln_conc_display,
                        int_display,
                        selected_wavelength
                    ])
                    # Store numeric values for sorting
                    item.setData(2, Qt.ItemDataRole.UserRole, float(soln_conc))
                    item.setData(3, Qt.ItemDataRole.UserRole, float(int_val))
                    # Store text for non-numeric columns
                    item.setData(0, Qt.ItemDataRole.UserRole, str(row.get('Solution Label', '')).lower())
                    item.setData(1, Qt.ItemDataRole.UserRole, self.current_element.lower())
                    item.setData(4, Qt.ItemDataRole.UserRole, selected_wavelength.lower())
                    items.append(item)

                for i, item in enumerate(items):
                    # Apply alternating background
                    if i % 2 == 0:
                        for col in range(5):
                            item.setBackground(col, QBrush(QColor("#fafafa")))
                self.details_tree.addTopLevelItems(items)
            except Exception as e:
                logger.error(f"Error populating tree: {str(e)}")
                item = CustomTreeWidgetItem([f"Error: {str(e)}", self.current_element, "", "", selected_wavelength])
                item.setForeground(0, QBrush(QColor("#d32f2f")))
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

        try:
            if 'Type' in df.columns:
                blk_data = df[df['Type'] == 'Blk']
            else:
                blk_data = df[df['Solution Label'].str.contains("BLANK", case=False, na=False)]
        except Exception as e:
            logger.error(f"Error filtering BLK data: {str(e)}")
            self.display_elements(["Cu", "Zn", "Fe"])
            return

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

        if self.filtered_elements:
            logger.info(f"Displaying filtered elements: {self.filtered_elements}")
            self.display_elements(self.filtered_elements)
        else:
            logger.info("No BLK elements found, displaying default elements")
            self.display_elements(["Cu", "Zn", "Fe"])