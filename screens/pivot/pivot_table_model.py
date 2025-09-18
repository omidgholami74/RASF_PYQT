from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt6.QtGui import QColor
from .oxide_factors import oxide_factors
import pandas as pd

class PivotTableModel(QAbstractTableModel):
    """Custom table model for pivot table, optimized for large datasets."""
    def __init__(self, pivot_tab, df=None, crm_rows=None):
        super().__init__()
        self.pivot_tab = pivot_tab
        self._df = df if df is not None else pd.DataFrame()
        self._crm_rows = crm_rows if crm_rows is not None else []
        self._row_info = []
        self._column_widths = {}
        self._build_row_info()

    def set_data(self, df, crm_rows=None):
        self.beginResetModel()
        self._df = df.copy()
        self._crm_rows = crm_rows if crm_rows is not None else []
        self._build_row_info()
        self.endResetModel()

    def _build_row_info(self):
        self._row_info = []
        for row_idx in range(len(self._df)):
            self._row_info.append({'type': 'pivot', 'index': row_idx})
            sol_label = self._df.iloc[row_idx]['Solution Label']
            for grp_idx, (sl, cdata) in enumerate(self._crm_rows):
                if sl == sol_label:
                    for sub in range(len(cdata)):
                        self._row_info.append({'type': 'crm', 'group': grp_idx, 'sub': sub})
                    break

    def rowCount(self, parent=QModelIndex()):
        return len(self._row_info)

    def columnCount(self, parent=QModelIndex()):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or index.row() >= len(self._row_info):
            return None

        row = index.row()
        col = index.column()
        col_name = self._df.columns[col]
        info = self._row_info[row]

        is_crm_row = False
        is_diff_row = False
        crm_row_data = None
        tags = None
        pivot_row = row

        if info['type'] == 'pivot':
            pivot_row = info['index']
        else:
            grp = info['group']
            sub = info['sub']
            _, crm_data = self._crm_rows[grp]
            if sub == 0:
                is_crm_row = True
                crm_row_data = crm_data[0][0]
                tags = crm_data[0][1]
            elif sub == 1:
                is_diff_row = True
                crm_row_data = crm_data[1][0]
                tags = crm_data[1][1]
            pivot_row = self._df.index[self._df['Solution Label'] == self._crm_rows[grp][0]].tolist()[0]

        if role == Qt.ItemDataRole.DisplayRole:
            dec = int(self.pivot_tab.decimal_places.currentText())
            if is_crm_row or is_diff_row:
                value = crm_row_data[col]
                return str(value) if value else ""
            else:
                value = self._df.iloc[pivot_row, col]
                if col_name != "Solution Label" and pd.notna(value):
                    try:
                        return f"{float(value):.{dec}f}"
                    except (ValueError, TypeError):
                        return "" if pd.isna(value) else str(value)
                return str(value) if pd.notna(value) else ""

        elif role == Qt.ItemDataRole.BackgroundRole:
            if is_crm_row:
                return QColor("#FFF5E4")
            elif is_diff_row and tags:
                if tags[col] == "in_range":
                    return QColor("#ECFFC4")
                elif tags[col] == "out_range":
                    return QColor("#FFCCCC")
                return QColor("#E6E6FA")
            return QColor("#f9f9f9") if pivot_row % 2 == 0 else QColor("white")

        elif role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignLeft if col_name == "Solution Label" else Qt.AlignmentFlag.AlignCenter

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._df.columns[section])
            return str(section + 1)
        return None

    def set_column_width(self, col, width):
        self._column_widths[col] = width