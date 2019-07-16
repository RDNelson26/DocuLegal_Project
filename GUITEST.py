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



global client_name
global doa
global attorney_name
global user_name
global user_email
global user_fax
global user_phone
global output_path
global widget
global lor_pl
global lop_pl
global rr_pl


#Path to files

path = os.getcwd()
settings_path = (path + "/settings")
lop_pl = (path + "\placeholders\LoP.docx")
rr_pl = (path + "\placeholders\RR.docx")
lor_pl = (path + "\placeholders\LoR.docx")

os.chdir (path+r"\gui")


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

def tabFocus(self):
    self.plainTextEdit.setTabChangesFocus(True)
    self.plainTextEdit_2.setTabChangesFocus(True)
    self.plainTextEdit_3.setTabChangesFocus(True)
    self.plainTextEdit_4.setTabChangesFocus(True)
    self.plainTextEdit_5.setTabChangesFocus(True)
    self.setTabOrder(self.plainTextEdit, self.plainTextEdit_2)
    self.setTabOrder(self.plainTextEdit_2, self.plainTextEdit_3)
    self.setTabOrder(self.plainTextEdit_3, self.plainTextEdit_4)
    self.setTabOrder(self.plainTextEdit_4, self.plainTextEdit_5)

provider_int = 0

#define user settings

if os.path.exists(settings_path + "/user_name.txt"):
    with open(settings_path + '/user_name.txt', 'r') as user_name_file:
        user_name = user_name_file.read()

if os.path.exists(settings_path + "/user_email.txt"):
    with open(settings_path + '/user_email.txt', 'r') as user_email_file:
        user_email = user_email_file.read()

if os.path.exists(settings_path + "/user_phone.txt"):
    with open(settings_path + '/user_phone.txt', 'r') as user_phone_file:
        user_phone = user_phone_file.read()

if os.path.exists(settings_path + "/user_fax.txt"):
    with open(settings_path + '/user_fax.txt', 'r') as user_fax_file:
        user_fax = user_fax_file.read()

if os.path.exists("/output_path.txt"):
    with open('/output_path.txt', 'r') as output_path_file:
        output_path = output_path_file.read()


def retrievesettings(settingsfile):
    with open(settingsfile, "r") as settingsoutput:
        settingsoutput.read()


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


#Create classes


#settings page

class settings(QDialog):
    def __init__(self):
        super(settings, self).__init__()
        loadUi("settings.ui", self)
        self.pushButton.clicked.connect(self.savesettings)
        self.pushButton_2.clicked.connect(self.movemain)
        tabFocus(self)

        if os.path.exists(settings_path + "/user_name.txt"):
            self.plainTextEdit.insertPlainText(user_name)

        if os.path.exists(settings_path + "/user_email.txt"):
            self.plainTextEdit_2.insertPlainText(user_email)

        if os.path.exists(settings_path + "/user_phone.txt"):
            self.plainTextEdit_3.insertPlainText(user_phone)

        if os.path.exists(settings_path + "/user_fax.txt"):
            self.plainTextEdit_4.insertPlainText(user_fax)

        if os.path.exists(settings_path + "/output_path.txt"):
            self.plainTextEdit_5.insertPlainText(output_path)

    def savesettings(self):

        user_name = self.plainTextEdit.toPlainText()
        self.writesettings("user_name", user_name)

        user_email = self.plainTextEdit_2.toPlainText()
        self.writesettings("user_email", user_email)

        user_phone = self.plainTextEdit_3.toPlainText()
        self.writesettings("user_phone", user_phone)

        user_fax = self.plainTextEdit_4.toPlainText()
        self.writesettings("user_fax", user_fax)


    def writesettings(self, settingname, settingstr):
        with open(settings_path + "/" + settingname + ".txt", "w") as settingfile:
            settingfile.write(settingstr)




    def movemain(self):
        sys.exit()

#provider number page

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


#LOP page

class LOPpage(QDialog):
    def __init__(self):
        super(LOPpage, self).__init__()
        loadUi("LOP.ui", self)
        self.pushButton.clicked.connect(self.create_click)
        self.pushButton_2.clicked.connect(self.cancel_click)

    def create_click(self):
        client_name = self.plainTextEdit.toPlainText()
        doa = self.plainTextEdit_2.toPlainText()
        attorney_name = self.plainTextEdit_3.toPlainText()
        print(client_name + " "  + doa + " " + attorney_name)

    def cancel_click(self):
        sys.exit()


#Reduction request page

class RRpage(QDialog):
    def __init__(self):
        super(RRpage, self).__init__()
        loadUi("RR.ui", self)
        self.pushButton.clicked.connect(self.create_click)
        self.pushButton_2.clicked.connect(self.cancel_click)

    def create_click(self):
        client_name = self.plainTextEdit.toPlainText()
        doa = self.plainTextEdit_2.toPlainText()
        attorney_name = self.plainTextEdit_3.toPlainText()


    def cancel_click(self):
        sys.exit()


#LOR page

class LORpage(QDialog):
    def __init__(self):
        super(LORpage, self).__init__()
        loadUi("LOR.ui", self)
        self.pushButton.clicked.connect(self.create_click)
        self.pushButton_2.clicked.connect(self.cancel_click)

    def create_click(self):
        client_name = self.plainTextEdit.toPlainText()
        doa = self.plainTextEdit_2.toPlainText()
        attorney_name = self.plainTextEdit_3.toPlainText()

        claim_num = self.plainTextEdit_4.toPlainText()
        def_insurance = self.plainTextEdit_5.toPlainText()
        def_adjuster_name = self.plainTextEdit_6.toPlainText()
        def_adjuster_address = self.plainTextEdit_7.toPlainText()
        def_adjuster_csz = self.plainTextEdit_8.toPlainText()
        def_adjuster_fax = self.plainTextEdit_9.toPlainText()

        doc = Document(lor_pl)

        if not os.path.exists(path + '/DocuLegal/LORs/Word'):
            os.makedirs(path + '/DocuLegal/LORs/Word')

        if not os.path.exists(path + '/DocuLegal/LORs/PDF'):
            os.makedirs(path + '/DocuLegal/LORs/PDF')

        replace_string(doc, client_name_PL, client_name)
        replace_string(doc, def_insurance_pl, def_insurance)
        replace_string(doc, claim_num_pl, claim_num)
        replace_string(doc, def_adjuster_name_pl, def_adjuster_name)
        replace_string(doc, def_adjuster_address_pl, def_adjuster_address)
        replace_string(doc, def_adjuster_csz_pl, def_adjuster_csz)
        replace_string(doc, def_adjuster_fax_pl, def_adjuster_fax)
        replace_string(doc, user_name_pl, user_name)
        replace_string(doc, user_number_pl, user_phone)
        replace_string(doc, attorney_name_PL, attorney_name)
        replace_string(doc, def_adjuster_name_pl, def_adjuster_name)


        doc.save('/DocuLegal/LORs/Word/' + client_name.upper() + ' LOR ' + def_insurance.upper() + ".docx")


    def cancel_click(self):
        sys.exit()


#Main page

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

if os.path.exists(settings_path + "/user_name.txt") and os.path.exists(settings_path + "/user_email.txt") and os.path.exists(settings_path + "/user_phone.txt") and os.path.exists(settings_path + "/user_email.txt"):
    app = QApplication(sys.argv)
    global widget
    widget = mainpage()
    widget.show()
    sys.exit(app.exec_())

else:
    app = QApplication(sys.argv)
    widget = settings()
    widget.show()
    sys.exit(app.exec())
