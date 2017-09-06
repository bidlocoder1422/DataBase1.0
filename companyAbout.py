# -*- coding: utf-8 -*-

#DONE1!1


from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd
class Ui_aboutCompanyWindow(object):
    def setupUi(self, aboutCompanyWindow,companyName):
        self.companyName=companyName
        aboutCompanyWindow.setObjectName("aboutCompanyWindow")
        aboutCompanyWindow.resize(490  , 260)
        self.companyNameHeaderLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.companyNameHeaderLabel.setGeometry(QtCore.QRect(190, 10, 371, 16))
        self.companyNameHeaderLabel.setText("")
        self.companyNameHeaderLabel.setObjectName("companyNameHeaderLabel")
        self.companyNameInfoLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.companyNameInfoLabel.setGeometry(QtCore.QRect(30, 40, 111, 16))
        self.companyNameInfoLabel.setObjectName("companyNameInfoLabel")
        self.companyNameOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.companyNameOutput.setGeometry(QtCore.QRect(150, 40, 301, 20))
        self.companyNameOutput.setObjectName("companyNameOutput")
        self.innInfoLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.innInfoLabel.setGeometry(QtCore.QRect(30, 70, 31, 16))
        self.innInfoLabel.setObjectName("innInfoLabel")
        self.kppInfoLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.kppInfoLabel.setGeometry(QtCore.QRect(30, 100, 41, 20))
        self.kppInfoLabel.setObjectName("kppInfoLabel")
        self.innOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.innOutput.setGeometry(QtCore.QRect(70, 70, 141, 20))
        self.innOutput.setObjectName("innOutput")
        self.kppOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.kppOutput.setGeometry(QtCore.QRect(70, 100, 141, 20))
        self.kppOutput.setObjectName("kppOutput")
        self.contractNumberLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.contractNumberLabel.setGeometry(QtCore.QRect(30, 130, 51, 16))
        self.contractNumberLabel.setObjectName("contractNumberLabel")
        self.contarctNumberOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.contarctNumberOutput.setGeometry(QtCore.QRect(80, 130, 161, 20))
        self.contarctNumberOutput.setObjectName("contarctNumberOutput")
        self.typeCompanyLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.typeCompanyLabel.setGeometry(QtCore.QRect(30, 160, 41, 16))
        self.typeCompanyLabel.setObjectName("typeCompanyLabel")
        self.typeCompanyLabelOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.typeCompanyLabelOutput.setGeometry(QtCore.QRect(60, 160, 111, 16))
        self.typeCompanyLabelOutput.setText("")
        self.typeCompanyLabelOutput.setObjectName("typeCompanyLabelOutput")
        self.contractNumberLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.contractNumberLabel.setGeometry(QtCore.QRect(30, 130, 51, 16))
        self.contractNumberLabel.setObjectName("contractNumberLabel")
        self.contarctNumberOutput = QtWidgets.QLineEdit(aboutCompanyWindow)
        self.contarctNumberOutput.setGeometry(QtCore.QRect(80, 130, 161, 20))
        self.contarctNumberOutput.setObjectName("contactNumberOutput")
        self.contactNameLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.contactNameLabel.setGeometry(QtCore.QRect(240, 70, 101, 16))
        self.contactNameLabel.setObjectName("contactNameLabel")
        self.phoneNumberLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.phoneNumberLabel.setGeometry(QtCore.QRect(240, 100, 31, 16))
        self.phoneNumberLabel.setObjectName("phoneNumberLabel")
        self.contactNameOutput = QtWidgets.QLabel(aboutCompanyWindow)
        self.contactNameOutput.setGeometry(QtCore.QRect(350, 70, 71, 16))
        self.contactNameOutput.setObjectName("contactNameOutput")
        self.phoneNumber = QtWidgets.QLabel(aboutCompanyWindow)
        self.phoneNumber.setGeometry(QtCore.QRect(280, 100, 191, 16))
        self.phoneNumber.setObjectName("phoneNumber")
        self.adressLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.adressLabel.setGeometry(QtCore.QRect(20, 190, 111, 16))
        self.adressLabel.setObjectName("adressLabel")
        self.adressOutput = QtWidgets.QLineEdit(aboutCompanyWindow)
        self.adressOutput.setGeometry(QtCore.QRect(150, 190, 321, 20))
        self.adressOutput.setObjectName("adressOutput")
        self.postAdressLabel = QtWidgets.QLabel(aboutCompanyWindow)
        self.postAdressLabel.setGeometry(QtCore.QRect(20, 220, 91, 16))
        self.postAdressLabel.setObjectName("postAdressLabel")
        self.postAdressOutput = QtWidgets.QLineEdit(aboutCompanyWindow)
        self.postAdressOutput.setGeometry(QtCore.QRect(120, 220, 351, 20))
        self.postAdressOutput.setObjectName("postAdressOutput")
        self.retranslateUi(aboutCompanyWindow)
        QtCore.QMetaObject.connectSlotsByName(aboutCompanyWindow)

    def retranslateUi(self, aboutCompanyWindow):
        _translate = QtCore.QCoreApplication.translate
        aboutCompanyWindow.setWindowTitle(_translate("aboutCompanyWindow", self.companyName))
        self.companyNameInfoLabel.setText(_translate("aboutCompanyWindow", "Название компании: "))
        self.innInfoLabel.setText(_translate("aboutCompanyWindow", "ИНН: "))
        self.kppInfoLabel.setText(_translate("aboutCompanyWindow", "КПП:"))
        self.contractNumberLabel.setText(_translate("aboutCompanyWindow", "Договор:"))
        self.typeCompanyLabel.setText(_translate("aboutCompanyWindow", "Тип:"))
        self.contactNameLabel.setText(_translate("aboutCompanyWindow", "Контактное лицо: "))
        self.phoneNumberLabel.setText(_translate("aboutCompanyWindow", "тел: "))
        self.adressLabel.setText(_translate("aboutCompanyWindow", "Юридический адрес"))
        self.postAdressLabel.setText(_translate("aboutCompanyWindow", "Почтовый адрес"))

    def getInfoFromFile(self, companyName):
        companiesListFile = xlrd.open_workbook('data/companiesBase.xlsx')
        workSheet = companiesListFile.sheet_by_index(0)
        companyInfo = []
        rowNumber = 0
        for i in range(1, workSheet.nrows):
            if workSheet.cell(i, 1).value == companyName:
                rowNumber = i
                continue
        for i in range(1, workSheet.ncols):
            companyInfo.append(workSheet.cell(rowNumber, i).value)
        for i in range(len(companyInfo)):
            if not companyInfo[i]:
                companyInfo[i]='Неизвестно'
        if not companyInfo[3]:
            companyInfo[3]='Нет договора'
        self.companyNameOutput.setText(companyInfo[0])
        self.innOutput.setText(str(int(companyInfo[1])))
        self.kppOutput.setText(str(int(companyInfo[2])))
        self.contarctNumberOutput.setText(str(companyInfo[3]))
        self.contactNameOutput.setText(companyInfo[6])
        self.phoneNumber.setText(companyInfo[7])
        self.adressOutput.setText(companyInfo[8])
        self.postAdressOutput.setText(companyInfo[9])
        if companyInfo[4]:
            self.companyNameHeaderLabel.setText('Оригинал договора получен')
            self.companyNameHeaderLabel.setStyleSheet('color: green')
        else:
            self.companyNameHeaderLabel.setText('Оригинал договора не получен')
            self.companyNameHeaderLabel.setStyleSheet('color: red')
        self.typeCompanyLabelOutput.setText(str(companyInfo[5]))