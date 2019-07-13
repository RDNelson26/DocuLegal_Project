import os
import sys
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog, QPushButton, QWidget, QMainWindow
from PyQt5.uic import loadUi

os.chdir (r"\Users\Ryan\PycharmProjects\untitled3\gui")

doc_type = "TEST"

class settings(QDialog):
    def __init__(self):
        super(settings, self).__init__()
        loadUi("settings.ui", self)
        self.pushButton.clicked.connect(self.savesettings)
        self.pushButton_2.clicked.connect(self.movemain)


    def savesettings(self):
        username = self.plainTextEdit.connect()
        useremail = self.plainTextEdit_2.connect()
        userphone = self.plainTextEdit_3.connect()
        userfax = self.plainTextEdit_4.connect()
        outputpath = self.plainTextEdit_5.connect()

        print("saved")
        self.hide()

    def movemain(self):
        self.hide()

class providerpage(QDialog):
    def __init__(self):
        super(providerpage, self).__init__()
        loadUi("providernumber.ui", self)
        self.pushButton.clicked.connect(self.move)
        self.LOPpage = LOPpage()
        self.RRpage = RRpage()

    def move(self):
        if  doc_type = "LOP":
            self.LOPpage.show()
            self.hide()
        if doc_type = "RR":
            self.RRpage.show()
            self.hide()

class LOPpage(QDialog):
    def __init__(self):
        super(LOPpage, self).__init__()
        loadUi("LOP.ui", self)



class LORpage(QDialog):
    def __init__(self):
        super(LORpage, self).__init__()
        loadUi("LOR.ui", self)

class RRpage(QDialog):
    def __init__(self):
        super(RRpage, self).__init__()
        loadUi("RR.ui", self)


class mainpage(QDialog):
    def __init__(self):
        super(mainpage, self).__init__()
        loadUi('mainpage.ui', self)
        self.pushButton.clicked.connect(self.moveLOP)
        self.pushButton_2.clicked.connect(self.moveLOR)
        self.pushButton_3.clicked.connect(self.moveRR)
        self.pushButton_4.clicked.connect(self.movesettings)
        self.settings = settings()
        self.providerpage = providerpage()
        self.LORpage = LORpage()

    def movesettings(self):
        self.settings.show()


    def moveLOR(self):
        self.LORpage.show()
        self.hide()
        global doc_type
        doc_type = "LOR"

    def moveLOP(self):
        self.providerpage.show()
        self.hide()
        global doc_type
        doc_type = "LOP"

    def moveRR(self):
        self.providerpage.show()
        self.hide()
        doc_type = "RR"



app = QApplication(sys.argv)
widget = mainpage()
widget.show()
sys.exit(app.exec_())

