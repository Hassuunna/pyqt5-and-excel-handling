import sys
from os import path
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType

formClass, baseClass = loadUiType(path.join(path.dirname(__file__),'design.ui'))

class MainApp(formClass, baseClass):
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)

def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()