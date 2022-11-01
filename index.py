import sys
from os import path
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.uic import loadUiType
import openpyxl
import pandas as pd

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

formClass, baseClass = loadUiType(path.join(path.dirname(__file__),'design3.ui'))

class MainApp(formClass, baseClass):
    start, end = 0, 0
    intervals = []
    df = None
    result_df = None
    final_result = pd.DataFrame()
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Browse_BTN.clicked.connect(self.addItemsToTable)
        self.Calculate_BTN.clicked.connect(self.Calculate)
        self.Append_BTN.clicked.connect(self.Append)
        self.Save_BTN.clicked.connect(self.Save)

    def addItemsToTable(self):
        myFileExcel = QFileDialog.getOpenFileName(self, '', '',"Excel files (*.xlsx *.xls)")[0]
        if myFileExcel == '':
            QMessageBox.about(self, "Error", "Please Choose a file")
            return
        self.df = pd.read_excel(myFileExcel)
        if self.df.size == 0:
            QMessageBox.about(self, "Warning", "file is empty")
            return
        self.df.fillna('', inplace=True)
        self.tableWidget.setRowCount(self.df.shape[0])
        self.tableWidget.setColumnCount(self.df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(self.df.columns)

        for row in self.df.iterrows():
            row_index = row[0]
            values = row[1]
            for col_index, myValue in enumerate(values):
                if isinstance(myValue, (float, int)):
                    myValue = int(myValue)
                mytableItem = QTableWidgetItem(str(myValue))
                self.tableWidget.setItem(row_index, col_index, mytableItem)


    def Calculate(self):
        #make this dynamic done
        first_header = self.df.columns.tolist()[0]
        mod_df = self.df.set_index(first_header)
        #what if startValue and endValue are empty? done
        self.start = int(self.startValue.text()) if not self.startValue.text() == '' else mod_df.index[0]
        self.end = int(self.endValue.text()) if not self.endValue.text() == '' else mod_df.index[-1]
        interval = mod_df.loc[self.start: self.end]
        # round done
        self.result_df = interval.aggregate(['min', 'max', 'mean']).astype(int)
        model = PandasModel(self.result_df)
        self.tableView.setModel(model)


    def Append(self):
        self.intrevals.append(str(self.start) + " - " + str(self.end))
        new_row = pd.Series({})
        self.final_result = pd.concat([self.final_result, new_row.to_frame().T, self.result_df])
        print(self.final_result)


    def Save (self):
        #append start and end
        #add merged big cell
        #don't overwrite
        myNewName = QFileDialog.getSaveFileName(self, 'Save as', '', 'Excel (*.xlsx)')[0]
        self.final_result.to_excel(myNewName, startcol=1)


def main():
    myApp = QApplication(sys.argv)
    myWindow = MainApp()
    myWindow.show()
    myApp.exec_()

if __name__ == '__main__':
    main()