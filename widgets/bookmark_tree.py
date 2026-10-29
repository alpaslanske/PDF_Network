from PySide6.QtWidgets import QTreeWidget
from lang import t

class BookmarkTree(QTreeWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderLabel(t("tree_header"))
        self.setAlternatingRowColors(True)
        self.setAnimated(True)
        self.setUniformRowHeights(True)
