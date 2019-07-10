import os
import sys
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog, QStackedLayout, QPushButton, QWidget, QMainWindow
from PyQt5.uic import loadUi, stupUi

os.chdir (r"\Users\Ryan\PycharmProjects\untitled3\gui")

class settings(QtWidgets.QMainWindow)
    def __init__(self, parent=None):
        super(settings, self).__init__()
        loadUi("settings.ui").self


class mainpage(QtWidgets.QMainWindow):
    def __init__(self):
        super(mainpage, self).__init__()
        self.startmainpage
        self.loadUi('mainpage.ui', self)
        self.pushButton_4.clicked.connect(self.move)
        self.settings = settings()
    def move(self):
        self.settings.show()




app = QApplication(sys.argv)
widget = mainpage()
widget.show()
sys.exit(app.exec_())

