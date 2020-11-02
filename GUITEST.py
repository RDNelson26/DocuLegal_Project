import os
import sys
import re
import time
from datetime import datetime
from docx import Document
import comtypes.client
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import Qt
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
global doc

#Path to files

path = os.getcwd()
settings_path = (path + "/settings")
lop_pl = (path + "\placeholders\LoP.docx")
rr_pl = (path + "\placeholders\RR.docx")
lor_pl = (path + "\placeholders\LoR.docx")
user_fax_pl = re.compile("User Fax")
user_email_pl = re.compile("User Email")
current_balance_pl = re.compile("Current Balance")
requested_balance_pl = re.compile("Requested Balance")

os.chdir (path+r"\gui")


#define replace_string , docxtopdf, and set values

def DocxtoPDF2(in_file, out_file):
    word = comtypes.client.CreateObject('Word.Application')
    d = word.Documents.Open(in_file)
    d.SaveAs(out_file, FileFormat=17)
    d.Close()
    word.Quit()

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
    self.textEdit.setTabChangesFocus(True)
    self.textEdit_2.setTabChangesFocus(True)
    self.textEdit_3.setTabChangesFocus(True)
    self.textEdit_4.setTabChangesFocus(True)
    self.textEdit_5.setTabChangesFocus(True)
    self.setTabOrder(self.textEdit, self.textEdit_2)
    self.setTabOrder(self.textEdit_2, self.textEdit_3)
    self.setTabOrder(self.textEdit_3, self.textEdit_4)
    self.setTabOrder(self.textEdit_4, self.textEdit_5)

#def tabFunc(self, x):
#    self.textEdit.setTabChangesFocus(True)
#    self.textEdit_2.setTabChangesFocus(True)
#    self.setTabOrder(self.textEdit, self.textEdit_2)
#    tabref=2
#
#    for tabref < x:
#        self.setTabOrder(self.textEdit_, self.textEdit_)


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
            self.textEdit.insertPlainText(user_name)

        if os.path.exists(settings_path + "/user_email.txt"):
            self.textEdit_2.insertPlainText(user_email)

        if os.path.exists(settings_path + "/user_phone.txt"):
            self.textEdit_3.insertPlainText(user_phone)

        if os.path.exists(settings_path + "/user_fax.txt"):
            self.textEdit_4.insertPlainText(user_fax)

        if os.path.exists(settings_path + "/output_path.txt"):
            self.textEdit_5.insertPlainText(output_path)

    def savesettings(self):

        user_name = self.textEdit.toPlainText()
        self.writesettings("user_name", user_name)

        user_email = self.textEdit_2.toPlainText()
        self.writesettings("user_email", user_email)

        user_phone = self.textEdit_3.toPlainText()
        self.writesettings("user_phone", user_phone)

        user_fax = self.textEdit_4.toPlainText()
        self.writesettings("user_fax", user_fax)


    def writesettings(self, settingname, settingstr):
        with open(settings_path + "/" + settingname + ".txt", "w") as settingfile:
            settingfile.write(settingstr)




    def movemain(self):
        self.close()

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
        client_name = self.textEdit.toPlainText()
        doa = self.textEdit_2.toPlainText()
        attorney_name = self.textEdit_3.toPlainText()
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
        client_name = self.textEdit.toPlainText()
        doa = self.textEdit_2.toPlainText()
        attorney_name = self.textEdit_3.toPlainText()


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
        client_name = self.textEdit.toPlainText()
        doa = self.textEdit_2.toPlainText()
        attorney_name = self.textEdit_3.toPlainText()
        claim_num = self.textEdit_4.toPlainText()
        def_insurance = self.textEdit_5.toPlainText()
        def_adjuster_name = self.textEdit_6.toPlainText()
        def_adjuster_address = self.textEdit_7.toPlainText()
        def_adjuster_csz = self.textEdit_8.toPlainText()
        def_adjuster_fax = self.textEdit_9.toPlainText()

        doc = Document(lor_pl)
        if not os.path.exists('C:/DocuLegal/LORs/Word'):
            os.makedirs('C:/DocuLegal/LORs/Word')

        if not os.path.exists('C:/DocuLegal/LORs/PDF'):
            os.makedirs('C:/DocuLegal/LORs/PDF')

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
        replace_string(doc, today_date_PL, today_date)


        doc.save('C:/DocuLegal/LORs/Word/' + client_name.upper() + ' LOR ' + def_insurance.upper() + ".docx")

        if str(self.comboBox.currentText()) == "Docx and PDF":

            DocxtoPDF2('C:/DocuLegal/LORs/Word/' + client_name.upper() + ' LOR ' + def_insurance.upper() + ".docx",
                      'C:/DocuLegal/LORs/PDF/' + client_name.upper() + ' LOR ' + def_insurance.upper() + ".pdf")

    def cancel_click(self):
        mainpage.show()
        self.hide()


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
        self.LOPpage = LOPpage()
        self.RRpage = RRpage()


    def movesettings(self):
        self.settings.show()

    def moveLOR(self):
        self.LORpage.show()
        self.hide()
        global doc_type
        doc_type = "LOR"

    def moveLOP(self):
        self.LOPpage.show()
        self.hide()
        global doc_type
        doc_type = "LOP"

    def moveRR(self):
        self.RRpage.show()
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
