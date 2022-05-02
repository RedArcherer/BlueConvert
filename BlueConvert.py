import sys

import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QGridLayout, QFileDialog, QLineEdit, \
    QWidget, QMessageBox, QListWidget, QLabel
from PyQt5.QtCore import Qt

import datetime
import os

class ListBoxWidget(QListWidget):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.resize(600, 600)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            self.links = []

            for url in event.mimeData().urls():
                self.links.append(str(url.toLocalFile()))
            self.addItems(self.links)

        else:
            event.ignore()

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setupUI()
        self.csvbrowse.clicked.connect(self.browseFiles)
        self.select.clicked.connect(self.Modify)

    def setupUI(self):
        # Add instructions to window and place correctly
        container = QWidget()

        self.instructions = QLabel("Please enter your CSV file. Use browse button or drag and drop file into the white box to do so.")
        self.instructions.setWordWrap(True)

        self.csvpath = QLineEdit()
        self.csvbrowse = QPushButton("Browse")

        self.dropBox = ListBoxWidget(self)
        self.select = QPushButton("Done")

        self.layout = QGridLayout(container)

        # Add csvpath, csvbrowse, dropbox, select, label

        self.layout.addWidget(self.instructions, 0, 0, 1, 4)
        self.layout.addWidget(self.dropBox, 1, 0, 2, 4)
        self.layout.addWidget(self.csvpath, 2, 0, 1, 3)
        self.layout.addWidget(self.csvbrowse, 2, 3)
        self.layout.addWidget(self.select, 3, 2)

        self.setCentralWidget(container)

        self.setGeometry(300, 300, 750, 250)
        self.setWindowTitle("Bluedart conversion")
        self.show()

    def browseFiles(self):
        self.fname = QFileDialog.getOpenFileName(self, 'Open File', 'C:', 'CSV file (*.csv)')
        self.csvpath.setText(self.fname[0])

    def Modify(self):
        try:
            self.csvpath.setText(self.dropBox.links[0])
        except:
            pass
        self.inputFileDF = pd.read_csv(self.csvpath.text())
        self.inputFileDF = self.inputFileDF.loc[:, ['Name', 'Shipping Name', 'Shipping Zip',
                                                    'Shipping Phone', 'Shipping Street', 'Total']]
        self.changeValues()

        self.ExportExcel()

        self.SuccessMessage()

    def ExportExcel(self):
        num = 1

        while len(self.outputDF.index) >= 25:
            # Copy row of outputDF to toExcel and drop it
            self.toExcel = self.outputDF.iloc[0:25]
            self.outputDF.drop(labels = [i for i in range((num - 1) * 25, num * 25)], axis = 0,
                               inplace = True)

            path = str(os.path.join(os.path.expanduser('~'), 'Desktop'))
            path = path.replace("\\", "/")
            path += "/"
            path += 'UPLOAD'
            path += str(num)
            path += ' '
            path += self.CreateDate()
            path += '.xlsx'
            self.toExcel.to_excel(path, index=False)

            self.toExcel = self.toExcel[0:0]
            num += 1

        if len(self.outputDF.index) != 0:
            path = str(os.path.join(os.path.expanduser('~'), 'Desktop'))
            path = path.replace("\\", "/")
            path += "/"
            path += 'UPLOAD'
            path += str(num)
            path += ' '
            path += self.CreateDate()
            path += '.xlsx'
            self.toExcel = self.outputDF.iloc[0:25]
            self.toExcel.to_excel(path, index = False)

    def changeValues(self):
        # If value needs to be changed, it is changed. Then the new value is placed directly into the
        # output dataframe

        date = self.CreateDate()
        date += ' '
        date += '12:00:00'
        self.outputDF = pd.read_excel("Bluedart Order template.xlsx")

        # Checking if row is null and deleting it if it is
        self.inputFileDF = self.inputFileDF.dropna(axis=0)

        for ind in range(0, len(self.inputFileDF.index)):

            # Altering Zip code
            zipcode = str(self.inputFileDF.iloc[ind, 2])
            zipcode = list(zipcode)
            # Checking if first character of zipcode is apostrophe and deleting apostrophe if it is
            if zipcode[0] == "'":
                zipcode.remove("'")
            zipcode = ''.join(zipcode)
            self.outputDF.at[ind, 'ConsigneePincode'] = zipcode

            # Altering phone numbers
            phone = self.inputFileDF.iloc[ind, 3]

            # Changing if phone number is a formula
            phonelist = list(str(phone))
            if 'E' in phonelist:
                e = phonelist.index('E')
                i = 0
                list1 = []

                while i < e:
                    list1.append(phonelist[i])
                    i += 1

                while phonelist[0] != '+':
                    phonelist.pop(0)
                phonelist.pop(0)

                eleven = ''.join(phonelist)
                num = ''.join(list1)

                num = float(num)
                eleven = int(eleven)

                phoneno = round(num * (pow(10, eleven)))
                phonelist = list(str(phoneno))

            # Making phone number proper 10 digits
            while (len(phonelist)) > 10:
                for x in phonelist:
                    try:
                        int(x)
                    except:
                        phonelist.remove(x)
                if phonelist[0] == '9' and phonelist[1] == '1':
                    phonelist.pop(0)
                    phonelist.pop(0)

                if phonelist[0] == '0':
                    phonelist.pop(0)

            phoneno = ''.join(phonelist)
            self.outputDF.at[ind, 'ConsigneeMobile'] = phoneno

            # Altering CRN to remove hashtag
            crn = self.inputFileDF.iloc[ind, 0]
            crn = list(crn)
            if crn[0] == '#':
                crn.pop(0)
            crn = ''.join(crn)
            self.outputDF.at[ind, 'CreditReferenceNo'] = crn

            # Copying the rest of the inds into outputDF
            self.outputDF.at[ind, 'ConsigneeName'] = self.inputFileDF.iloc[ind, 1]
            self.outputDF.at[ind, 'ConsigneeAttention'] = self.inputFileDF.iloc[ind, 1]
            self.outputDF.at[ind, 'ConsigneeAddress1'] = self.inputFileDF.iloc[ind, 4]
            self.outputDF.at[ind, 'ProductCode'] = 'D'
            self.outputDF.at[ind, 'ProductType'] = 'NDOX'
            self.outputDF.at[ind, 'PieceCount'] = '1'
            self.outputDF.at[ind, 'DeclaredValue'] = self.inputFileDF.iloc[ind, 5]
            self.outputDF.at[ind, 'InvoiceNo'] = '0'
            self.outputDF.at[ind, 'PickupDate'] = date
            self.outputDF.at[ind, 'PickupTime'] = '1600'

            # Adding
            self.outputDF.at[ind, 'OriginArea'] = 'IMP'
            self.outputDF.at[ind, 'CustomerCode'] = '000206'
            self.outputDF.at[ind, 'CustomerName'] = 'MADAKE BAMBOO SOLUTIONS LLP'
            self.outputDF.at[ind, 'CustomerAddress1'] = 'Kasturi Building'
            self.outputDF.at[ind, 'CustomerAddress2'] = 'Thangal Bazar'
            self.outputDF.at[ind, 'CustomerAddress3'] = 'Imphal'
            self.outputDF.at[ind, 'CustomerPincode'] = '795001'
            self.outputDF.at[ind, 'CustomerTelephone'] = '000206'
            self.outputDF.at[ind, 'CustomerTelephone'] = '6374679609'
            self.outputDF.at[ind, 'CustomerMobile'] = '6374679609'
            self.outputDF.at[ind, 'Sender'] = 'MADAKE BAMBOO SOLUTIONS LLP'
            self.outputDF.at[ind, 'IsToPayCustomer'] = 'FALSE'

    def CreateDate(self):
        date = datetime.datetime.now()

        day = int(date.day)

        tempday = day / 10
        if tempday < 1:
            day = '0'
            day += str(date.day)
        else:
            day = str(date.day)

        month = int(date.month)
        tempmonth = month / 10

        if tempmonth < 1:
            month = '0'
            month += str(date.month)
        else:
            month = str(date.month)

        finaldate = ''
        finaldate += day
        finaldate += '-'
        finaldate += month
        finaldate += '-'
        finaldate += str(date.year)

        return finaldate

    def SuccessMessage(self):
        msg = QMessageBox()
        msg.setWindowTitle('Success')
        msg.setText('Conversion Successful!')
        msg.setInformativeText('Check your desktop to access the Excel files')
        msg.setIcon(QMessageBox.Information)
        x = msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()

    sys.exit(app.exec_())
