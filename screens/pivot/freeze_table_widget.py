from PyQt6.QtWidgets import QTableView, QAbstractItemView,QHeaderView
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

class FreezeTableWidget(QTableView):
    """A QTableView with a frozen first column that does not scroll horizontally."""
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.frozenTableView = QTableView(self)
        self.setModel(model)
        self.frozenTableView.setModel(model)
        self.init()

        self.horizontalHeader().sectionResized.connect(self.updateSectionWidth)
        self.verticalHeader().sectionResized.connect(self.updateSectionHeight)
        self.frozenTableView.verticalScrollBar().valueChanged.connect(self.frozenVerticalScroll)
        self.verticalScrollBar().valueChanged.connect(self.mainVerticalScroll)

    def init(self):
        self.frozenTableView.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.frozenTableView.verticalHeader().hide()
        self.frozenTableView.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.viewport().stackUnder(self.frozenTableView)
        
        self.frozenTableView.setStyleSheet("""
            QTableView { 
                border: none;
                selection-background-color: #999;
            }
        """)
        self.frozenTableView.setSelectionModel(self.selectionModel())
        
        for col in range(self.model().columnCount()):
            self.frozenTableView.setColumnHidden(col, col != 0)
        
        self.frozenTableView.setColumnWidth(0, self.columnWidth(0))
        self.frozenTableView.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.frozenTableView.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.frozenTableView.show()
        self.frozenTableView.viewport().repaint()
        
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        self.frozenTableView.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        
        self.updateFrozenTableGeometry()

    def updateSectionWidth(self, logicalIndex, oldSize, newSize):
        if logicalIndex == 0:
            self.frozenTableView.setColumnWidth(0, newSize)
            self.updateFrozenTableGeometry()
            self.frozenTableView.viewport().repaint()

    def updateSectionHeight(self, logicalIndex, oldSize, newSize):
        self.frozenTableView.setRowHeight(logicalIndex, newSize)
        self.frozenTableView.viewport().repaint()

    def frozenVerticalScroll(self, value):
        self.viewport().stackUnder(self.frozenTableView)
        self.verticalScrollBar().setValue(value)
        self.frozenTableView.viewport().repaint()
        self.viewport().update()

    def mainVerticalScroll(self, value):
        self.viewport().stackUnder(self.frozenTableView)
        self.frozenTableView.verticalScrollBar().setValue(value)
        self.frozenTableView.viewport().repaint()
        self.viewport().update()

    def updateFrozenTableGeometry(self):
        self.frozenTableView.setGeometry(
            self.verticalHeader().width() + self.frameWidth(),
            self.frameWidth(),
            self.columnWidth(0),
            self.viewport().height() + self.horizontalHeader().height()
        )
        self.frozenTableView.viewport().repaint()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.updateFrozenTableGeometry()
        self.frozenTableView.viewport().repaint()

    def moveCursor(self, cursorAction, modifiers):
        current = super().moveCursor(cursorAction, modifiers)
        if cursorAction == QAbstractItemView.CursorAction.MoveLeft and current.column() > 0:
            visual_x = self.visualRect(current).topLeft().x()
            if visual_x < self.frozenTableView.columnWidth(0):
                new_value = self.horizontalScrollBar().value() + visual_x - self.frozenTableView.columnWidth(0)
                self.horizontalScrollBar().setValue(int(new_value))
        self.frozenTableView.viewport().repaint()
        return current

    def scrollTo(self, index, hint=QAbstractItemView.ScrollHint.EnsureVisible):
        if index.column() > 0:
            super().scrollTo(index, hint)
        self.frozenTableView.viewport().repaint()