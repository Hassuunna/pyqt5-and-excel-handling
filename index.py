import sys
from os import path
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
import pandas as pd

formClass, baseClass = loadUiType(path.join(path.dirname(__file__),'design.ui'))

class MainApp(formClass, baseClass):
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Browse_BTN.clicked.connect(self.addItemsToTable)
        self.Calculate_BTN.clicked.connect(self.Average)
        self.Save_BTN.clicked.connect(self.Save)

    def addItemsToTable(self):
        fileName = QFileDialog.getOpenFileName(self, 'OpenFile')[0]
        df = pd.read_excel(fileName)
        if df.size == 0:
            return
        df.fillna('', inplace=True)
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)
        # returns pandas array object
        for row in df.iterrows():
            values = row[1]
            for col_index, value in enumerate(values):
                if isinstance(value, (float, int)):
                    value = '{0:0,.0f}'.format(value)
                tableItem = QTableWidgetItem(str(value))
                self.tableWidget.setItem(row[0], col_index, tableItem)

    def Average(self):
        #column = 0
        numberofCurrentRows = self.tableWidget.rowCount()
        numberofCurrentColumns = self.tableWidget.columnCount()
        averages = [0] * numberofCurrentColumns
        # rowCount() This property holds the number of rows in the table
        for row in range(numberofCurrentRows):
            for column in range(numberofCurrentColumns):
                # item(row, 0) Returns the item for the given row and column if one has been set; otherwise returns nullptr.
                _item = self.tableWidget.item(row, column) 
                if _item:
                    item = self.tableWidget.item(row, column).text().replace(',','')
                    averages[column] += float(item)
        averages2 = [x / numberofCurrentRows for x in averages]
        self.tableWidget.insertRow(numberofCurrentRows)
        for column in range(numberofCurrentColumns):
            tableItem = QTableWidgetItem(str(averages2[column]))
            self.tableWidget.setItem(numberofCurrentRows, column,tableItem)
    

def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()