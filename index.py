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
        myNumberofCurrentRows = self.tableWidget.rowCount()
        myNumberofCurrentColumns = self.tableWidget.columnCount()
        mySums = [0] * myNumberofCurrentColumns
        
        for row in range(self.myAverageRow):
            for column in range(myNumberofCurrentColumns):
                myItem = self.tableWidget.item(row, column)
                if myItem:
                    myItemText = myItem.text().replace(',','')
                    mySums[column] += float(myItemText)
        
        myAverages = [x / self.myAverageRow for x in mySums]
        if self.myAverageRow >= myNumberofCurrentRows:
            self.tableWidget.insertRow(self.myAverageRow)
        for column in range(myNumberofCurrentColumns):
            tableItem = QTableWidgetItem(str(myAverages[column]))
            self.tableWidget.setItem(self.myAverageRow, column,tableItem)
    
    def Save(self):
        columnHeaders = []

        # create column header list
        for j in range(self.tableWidget.model().columnCount()):
            columnHeaders.append(self.tableWidget.horizontalHeaderItem(j).text())

        df = pd.DataFrame(columns=columnHeaders)

        # create dataframe object recordset
        for row in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                df.at[row, columnHeaders[col]] = self.tableWidget.item(row, col).text()

        df.to_excel('Dummy.xlsx', index=False)
        print('Excel file exported')


def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()