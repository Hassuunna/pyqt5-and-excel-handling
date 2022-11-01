import sys
from os import path
from unicodedata import name
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from numpy import Infinity
import pandas as pd

formClass, baseClass = loadUiType(path.join(path.dirname(__file__),'design.ui'))

class MainApp(formClass, baseClass):
    df = None
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Browse_BTN.clicked.connect(self.addItemsToTable)
        self.Calculate_BTN.clicked.connect(self.Average)
        self.Save_BTN.clicked.connect(self.Save)

    def addItemsToTable(self):
        myFileExcel = QFileDialog.getOpenFileName(self, '', '',"Excel files (*.xlsx)")[0]
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
                    myValue = '{0:.0f}'.format(myValue)
                mytableItem = QTableWidgetItem(str(myValue))
                self.tableWidget.setItem(row_index, col_index, mytableItem)


    def Average(self):
        myCurrentRows = self.df.shape[0]
        myCurrentColumns = self.df.shape[1]
        mySums = [0] * myCurrentColumns
        myMax = [-Infinity] * myCurrentColumns
        myMin = [Infinity] * myCurrentColumns
        
        for row in range(myCurrentRows):
            for column in range(myCurrentColumns):
                myItem = self.tableWidget.item(row, column)
                if myItem:
                    myItemText = float(myItem.text())
                    mySums[column] += myItemText
                    if myItemText > myMax[column]:
                        myMax[column] = myItemText
                    if myItemText < myMin[column]:
                        myMin[column] = myItemText

        myAverages = [x / myCurrentRows for x in mySums]

        #self.addRow(myMin)
        self.tableWidget.insertRow(myCurrentRows)
        for column in range(myCurrentColumns):
            tableItem = QTableWidgetItem(str(myMin[column]))
            self.tableWidget.setItem(myCurrentRows, column,tableItem)

        self.tableWidget.insertRow(myCurrentRows+1)
        for column in range(myCurrentColumns):
            tableItem = QTableWidgetItem(str(myMax[column]))
            self.tableWidget.setItem(myCurrentRows+1, column,tableItem)

        self.tableWidget.insertRow(myCurrentRows+2)
        for column in range(myCurrentColumns):
            tableItem = QTableWidgetItem(str(myAverages[column]))
            self.tableWidget.setItem(myCurrentRows+2, column,tableItem)

    def addRow(self, listToAdd):
        #addlist
        print("row added succesfully!")


    def Save (self):
        columnHeaders = self.df.columns.tolist()
        df = pd.DataFrame(columns=columnHeaders)

        for row in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                df.at[row, columnHeaders[col]] = self.tableWidget.item(row, col).text()

        myNewName = QFileDialog.getSaveFileName(self,'Save as', '', 'Excel (*.xlsx)')[0]
        
        df.to_excel(myNewName)


def main():
    myApp = QApplication(sys.argv)
    myWindow = MainApp()
    myWindow.show()
    myApp.exec_()

if __name__ == '__main__':
    main()