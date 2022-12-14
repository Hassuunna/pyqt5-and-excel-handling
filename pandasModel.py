from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class PandasModel(QAbstractTableModel):
    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data
    
    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return QVariant(str(
                    self._data.iloc[index.row()][index.column()]))
        return QVariant()

    def headerData(self, rowcol, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[rowcol]
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return self._data.index[rowcol]
