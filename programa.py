from openpyxl import load_workbook
import xlsxwriter
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
import threading

class Ui_MainPrincipal(QWidget):
    def setupUi(self, MainPrincipal):
        MainPrincipal.setObjectName("MainPrincipal")
        MainPrincipal.resize(469, 231)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainPrincipal.sizePolicy().hasHeightForWidth())
        MainPrincipal.setSizePolicy(sizePolicy)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainPrincipal = QtWidgets.QMainWindow()
    ui = Ui_MainPrincipal()
    ui.setupUi(MainPrincipal)

    MainPrincipal.show()
    
    sys.exit(app.exec_())