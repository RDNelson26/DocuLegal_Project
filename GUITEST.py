import os
import sys
import re
import time
from datetime import datetime
from docx import Document
import comtypes.client
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog, QPushButton, QWidget, QMainWindow
from PyQt5.uic import loadUi



#Path to files

path = os.getcwd()
lop_pl = (path + "\placeholders\LoP.docx")
rr_pl = (path + "\placeholders\RR.docx")
lor_pl = (path + "\placeholders\LoR.docx")



#define replace_string , docxtopdf, and set values

def DocxtoPDF(inputFileName, outputFileName, formatType = 17):
    Word = comtypes.client.CreateObject("Word.Application")
    Word.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = Word.Documents.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)
    deck.Close()
    Word.Quit()



def replace_string(doc_obj, regex, replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_string(cell, regex, replace)


provider_int = 0

#Placeholder values

#LoP and universal

provider_email_PL = re.compile("Provider Email")
provider_name_PL = re.compile("Provider Name")
client_name_PL = re.compile("Client Name")
doa_PL = re.compile("Date of Accident")
today_date_PL = re.compile("Today Date")
provider_address_PL = re.compile("Provider Address")
lop_amount_PL = re.compile("LOP Amount")
attorney_name_PL = re.compile("Attorney Name")

#LoR

def_insurance_pl = re.compile("Defendant Insurance")
claim_num_pl = re.compile("CNumber")
def_adjuster_name_pl = re.compile("AdjusterName")
def_adjuster_address_pl = re.compile("Defendant Adjuster Address")
def_adjuster_csz_pl = re.compile("Defendant Adjuster CSZ")
def_adjuster_fax_pl = re.compile("Defendant Adjuster Fax")
user_name_pl = re.compile("User Name")
user_number_pl = re.compile("User Number")

#RR

user_fax_pl = re.compile("User Fax")
user_email_pl = re.compile("User Email")
current_balance_pl = re.compile("Current Balance")
requested_balance_pl = re.compile("Requested Balance")

#Create date
now = datetime.now()
today_date = now.strftime('%m/%d/%Y')

os.chdir (path+r"\gui")


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
        if  doc_type == "LOP":
            self.LOPpage.show()
            self.hide()

        if doc_type == "RR":
            self.RRpage.show()
            self.hide()

class LOPpage(QDialog):
    def __init__(self):
        super(LOPpage, self).__init__()
        loadUi("LOP.ui", self)

class RRpage(QDialog):
    def __init__(self):
        super(RRpage, self).__init__()
        loadUi("RR.ui", self)


class LORpage(QDialog):
    def __init__(self):
        super(LORpage, self).__init__()
        loadUi("LOR.ui", self)



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
        global doc_type
        doc_type = "RR"



app = QApplication(sys.argv)
widget = mainpage()
widget.show()
sys.exit(app.exec_())

