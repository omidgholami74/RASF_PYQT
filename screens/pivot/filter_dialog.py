from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QTreeView, QPushButton
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem

class FilterDialog(QDialog):
    """Dialog for row/column filtering."""
    def __init__(self, parent, title, is_row_filter=True):
        super().__init__(parent)
        self.parent = parent
        self.is_row_filter = is_row_filter
        self.setWindowTitle(title)
        self.setGeometry(200, 200, 300, 460)
        self.setModal(True)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select Field:"))
        
        self.field_combo = QComboBox()
        if self.is_row_filter:
            self.field_combo.addItems(["Solution Label", "Element"])
        else:
            self.field_combo.addItems(["Element"])
        layout.addWidget(self.field_combo)

        self.tree_view = QTreeView()
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["Value", "Select"])
        self.tree_view.setModel(self.model)
        self.tree_view.setRootIsDecorated(False)
        self.tree_view.header().setStretchLastSection(False)
        self.tree_view.header().resizeSection(0, 160)
        self.tree_view.header().resizeSection(1, 80)
        layout.addWidget(self.tree_view)

        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(lambda: self.set_all_checks(True))
        button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(lambda: self.set_all_checks(False))
        button_layout.addWidget(deselect_all_btn)
        
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.apply_filters)
        button_layout.addWidget(apply_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        self.field_combo.currentTextChanged.connect(self.update_tree)
        self.update_tree()

    def update_tree(self):
        self.model.clear()
        self.model.setHorizontalHeaderLabels(["Value", "Select"])
        field = self.field_combo.currentText()
        if not field:
            return
        
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        uniques = (self.parent.solution_label_order if self.is_row_filter and field == "Solution Label" else
                  sorted(self.parent.pivot_data[field].astype(str).unique()) if field in self.parent.pivot_data.columns else
                  self.parent.element_order)
        
        if field not in var_dict:
            var_dict[field] = {v: True for v in uniques}
        
        for val in uniques:
            value_item = QStandardItem(str(val))
            check_item = QStandardItem()
            check_item.setCheckable(True)
            check_item.setCheckState(Qt.CheckState.Checked if var_dict[field].get(val, True) else Qt.CheckState.Unchecked)
            self.model.appendRow([value_item, check_item])
        
        self.tree_view.clicked.connect(self.toggle_check)

    def toggle_check(self, index):
        if index.column() != 1:
            return
        val = self.model.item(index.row(), 0).text()
        field = self.field_combo.currentText()
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        var_dict[field][val] = not var_dict[field].get(val, True)
        self.model.item(index.row(), 1).setCheckState(
            Qt.CheckState.Checked if var_dict[field][val] else Qt.CheckState.Unchecked)

    def set_all_checks(self, value):
        field = self.field_combo.currentText()
        var_dict = self.parent.row_filter_values if self.is_row_filter else self.parent.column_filter_values
        if field not in var_dict:
            return
        for val in var_dict[field]:
            var_dict[field][val] = value
        self.update_tree()

    def apply_filters(self):
        self.parent.update_pivot_display()
        self.accept()