# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\addcompany.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import openpyxl,xlwt,xlrd


class addCompanyDialog(QtWidgets.QWidget):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(468, 324)
        self.companyNameInput = QtWidgets.QLineEdit(Dialog)
        self.companyNameInput.setGeometry(QtCore.QRect(79, 40, 341, 20))
        self.companyNameInput.setObjectName("companyNameInput")
        self.companyNameLbl = QtWidgets.QLabel(Dialog)
        self.companyNameLbl.setGeometry(QtCore.QRect(7, 43, 71, 16))
        self.companyNameLbl.setObjectName("companyNameLbl")
        self.innLabel = QtWidgets.QLabel(Dialog)
        self.innLabel.setGeometry(QtCore.QRect(16, 74, 47, 13))
        self.innLabel.setObjectName("innLabel")
        self.innInput = QtWidgets.QLineEdit(Dialog)
        self.innInput.setGeometry(QtCore.QRect(80, 70, 131, 20))
        self.innInput.setObjectName("innInput")
        self.kppLabel = QtWidgets.QLabel(Dialog)
        self.kppLabel.setGeometry(QtCore.QRect(250, 75, 47, 13))
        self.kppLabel.setObjectName("kppLabel")
        self.kppInput = QtWidgets.QLineEdit(Dialog)
        self.kppInput.setGeometry(QtCore.QRect(290, 70, 131, 20))
        self.kppInput.setObjectName("kppInput")
        self.contractNumberInput = QtWidgets.QLineEdit(Dialog)
        self.contractNumberInput.setGeometry(QtCore.QRect(90, 110, 331, 20))
        self.contractNumberInput.setObjectName("contractNumberInput")
        self.contractLabel = QtWidgets.QLabel(Dialog)
        self.contractLabel.setGeometry(QtCore.QRect(10, 110, 47, 13))
        self.contractLabel.setObjectName("contractLabel")
        self.originalContractCheckBox = QtWidgets.QCheckBox(Dialog)
        self.originalContractCheckBox.setGeometry(QtCore.QRect(300, 140, 121, 17))
        self.originalContractCheckBox.setObjectName("originalContractCheckBox")
        self.companyTypeCB = QtWidgets.QComboBox(Dialog)
        self.companyTypeCB.setGeometry(QtCore.QRect(20, 140, 141, 22))
        self.companyTypeCB.setObjectName("companyTypeCB")
        self.headerLabel = QtWidgets.QLabel(Dialog)
        self.headerLabel.setGeometry(QtCore.QRect(130, 10, 121, 16))
        self.headerLabel.setObjectName("headerLabel")
        self.acceptBtn=QtWidgets.QPushButton(Dialog)
        self.acceptBtn.setObjectName("acceptBtn")
        self.acceptBtn.setGeometry(190,290,75,23)
        self.acceptBtn.setText('Принять')
        self.contactNamelabel = QtWidgets.QLabel(Dialog)
        self.contactNamelabel.setGeometry(QtCore.QRect(10, 180, 111, 16))
        self.contactNamelabel.setObjectName("contactNamelabel")
        self.contactNameIutput = QtWidgets.QLineEdit(Dialog)
        self.contactNameIutput.setGeometry(QtCore.QRect(123, 179, 113, 20))
        self.contactNameIutput.setObjectName("contactNameIutput")
        self.phoneLabel = QtWidgets.QLabel(Dialog)
        self.phoneLabel.setGeometry(QtCore.QRect(250, 180, 31, 16))
        self.phoneLabel.setObjectName("phoneLabel")
        self.phoneNumberInput = QtWidgets.QLineEdit(Dialog)
        self.phoneNumberInput.setGeometry(QtCore.QRect(290, 180, 171, 20))
        self.phoneNumberInput.setObjectName("phoneNumberInput")
        self.adressLabel = QtWidgets.QLabel(Dialog)
        self.adressLabel.setGeometry(QtCore.QRect(10, 220, 111, 16))
        self.adressLabel.setObjectName("adressLabel")
        self.adressInput = QtWidgets.QLineEdit(Dialog)
        self.adressInput.setGeometry(QtCore.QRect(130, 220, 331, 20))
        self.adressInput.setObjectName("adressInput")
        self.postAdressLabel = QtWidgets.QLabel(Dialog)
        self.postAdressLabel.setGeometry(QtCore.QRect(10, 260, 81, 16))
        self.postAdressLabel.setObjectName("postAdressLabel")
        self.postAdressInput = QtWidgets.QLineEdit(Dialog)
        self.postAdressInput.setGeometry(QtCore.QRect(100, 260, 361, 20))
        self.postAdressInput.setObjectName("postAdressInput")
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Новый контрагент"))
        self.companyNameLbl.setText(_translate("Dialog", "Имя компании"))
        self.innLabel.setText(_translate("Dialog", "ИНН"))
        self.kppLabel.setText(_translate("Dialog", "КПП"))
        self.contractLabel.setText(_translate("Dialog", "Договор:"))
        self.originalContractCheckBox.setText(_translate("Dialog", "Получен оригинал"))
        self.headerLabel.setText(_translate("Dialog", "Новый Контрагент"))
        self.acceptBtn.clicked.connect(self.acceptbtnPushed)
        self.companyTypeCB.clear()
        self.companyTypeCB.addItems(['Покупатель','Поставщик'])
        self.contactNamelabel.setText(_translate("Dialog", "Конатктное лицо : "))
        self.phoneLabel.setText(_translate("Dialog", "тел"))
        self.acceptBtn.setText(_translate("Dialog", "Принять"))
        self.adressLabel.setText(_translate("Dialog", "Юридический  адрес"))
        self.postAdressLabel.setText(_translate("Dialog", "Почтоый адрес"))

    @QtCore.pyqtSlot()
    def acceptbtnPushed(self):
        companyName=self.companyNameInput.text()
        companyINN=self.innInput.text()
        companyKPP=self.kppInput.text()
        companyType=self.companyTypeCB.currentText()
        contractNumber=self.contractNumberInput.text()
        if not contractNumber:
            contractNumber='Нет договора'
        checkOriginal=self.originalContractCheckBox.isChecked()
        contactName=self.contactNameIutput.text()
        phoneNUmber=self.phoneNumberInput.text()
        companyAdress=self.adressInput.text()
        postAdress=self.postAdressInput.text()
        excelFile=xlrd.open_workbook('data\companiesBase.xlsx')
        workSheet=excelFile.sheet_by_index(0)
        maxRow=workSheet.nrows
        excelFile=openpyxl.load_workbook('data\companiesBase.xlsx')
        workSheet=excelFile.get_sheet_by_name('sheet1')
        maxRow=maxRow+1
        maxRow=str(maxRow)
        workSheet['A'+maxRow]=maxRow
        workSheet['B'+maxRow]=companyName
        workSheet['C'+maxRow]=companyINN
        workSheet['D'+maxRow]=companyKPP
        workSheet['E' + maxRow]=contractNumber
        workSheet['F' + maxRow]=int(checkOriginal)
        workSheet['G' + maxRow]=companyType
        workSheet['H' + maxRow] = contactName
        workSheet['I' + maxRow] = phoneNUmber
        workSheet['J'+maxRow]=postAdress
        workSheet['K'+maxRow]=companyAdress
        excelFile.save('data\companiesBase.xlsx')
        self.acceptBtn.setVisible(0)