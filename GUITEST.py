import os
import sys
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.uic import loadUi




wd = os.getcwd()
os.chdir(wd + ('\gui'))



class mainpage(QDialog):
    def __init__(self):
        super(mainpage, self).__init__()
        loadUi('mainpage.ui', self)
        self.pushButton.clicked.connect(self.settings)

class settings(QDialog):
    def __init__(self):
        super(settings, self).__init__()
        loadUi('settings.ui', self)

class providernumber(QDialog):
    def __init__(self):
        super(providernumber, self).__init__()
        loadUi('providernumber.ui', self)

class LOR(QDialog):
    def __init__(self):
        super(LOR, self).__init__()
        loadUi('LOR.ui', self)

class LOP(QDialog):
    def __init__(self):
        super(LOP, self).__init__()
        loadUi('LOP.ui', self)

class RR(QDialog):
    def __init__(self):
        super(RR, self).__init__()
        loadUi('RR.ui', self)

app = QApplication(sys.argv)
widget = mainpage()
widget.show()
sys.exit(app.exec_())


