# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\billcard.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import sys, xlrd,math

class Ui_bilCard(QtWidgets.QWidget):

    def setupUi(self, bilCard, billData):
        bilCard.setObjectName("bilCard")
        bilCard.resize(640, 224)
        self.headerLabel = QtWidgets.QLabel(bilCard)
        self.headerLabel.setGeometry(QtCore.QRect(230, 10, 161, 16))
        self.headerLabel.setObjectName("headerLabel")
        self.companyNameLabel = QtWidgets.QLabel(bilCard)
        self.companyNameLabel.setGeometry(QtCore.QRect(30, 60, 121, 16))
        self.companyNameLabel.setObjectName("companyNameLabel")
        self.companyNameOutput = QtWidgets.QLabel(bilCard)
        self.companyNameOutput.setGeometry(QtCore.QRect(140, 60, 401, 16))
        self.companyNameOutput.setText("")
        self.companyNameOutput.setObjectName("companyNameOutput")
        self.priceLabel = QtWidgets.QLabel(bilCard)
        self.priceLabel.setGeometry(QtCore.QRect(30, 100, 51, 16))
        self.priceLabel.setObjectName("priceLabel")
        self.priceOutput = QtWidgets.QLabel(bilCard)
        self.priceOutput.setGeometry(QtCore.QRect(80, 100, 121, 16))
        self.priceOutput.setText("")
        self.priceOutput.setObjectName("priceOutput")
        self.deliveryYesNoOutput = QtWidgets.QLabel(bilCard)
        self.deliveryYesNoOutput.setGeometry(QtCore.QRect(220, 130, 141, 20))
        self.deliveryYesNoOutput.setText("")
        self.deliveryYesNoOutput.setObjectName("deliveryYesNoOutput")
        self.documentTypeLabel = QtWidgets.QLabel(bilCard)
        self.documentTypeLabel.setGeometry(QtCore.QRect(30, 160, 121, 16))
        self.documentTypeLabel.setObjectName("documentTypeLabel")
        self.documentTypeOutput = QtWidgets.QLabel(bilCard)
        self.documentTypeOutput.setGeometry(QtCore.QRect(113, 160, 47, 16))
        self.documentTypeOutput.setText("")
        self.documentTypeOutput.setObjectName("documentTypeOutput")
        self.documentNumberLabel = QtWidgets.QLabel(bilCard)
        self.documentNumberLabel.setGeometry(QtCore.QRect(30, 190, 91, 16))
        self.documentNumberLabel.setObjectName("documentNumberLabel")
        self.documentNumberOutput = QtWidgets.QLabel(bilCard)
        self.documentNumberOutput.setGeometry(QtCore.QRect(124, 190, 81, 16))
        self.documentNumberOutput.setText("")
        self.documentNumberOutput.setObjectName("documentNumberOutput")
        self.numberOfPositionsLabel = QtWidgets.QLabel(bilCard)
        self.numberOfPositionsLabel.setGeometry(QtCore.QRect(270, 100, 111, 16))
        self.numberOfPositionsLabel.setObjectName("numberOfPositionsLabel")
        self.bilTypeLabel = QtWidgets.QLabel(bilCard)
        self.bilTypeLabel.setGeometry(QtCore.QRect(270, 190, 61, 16))
        self.bilTypeLabel.setObjectName("bilTypeLabel")
        self.bilTypeOutput = QtWidgets.QLabel(bilCard)
        self.bilTypeOutput.setGeometry(QtCore.QRect(333, 190, 61, 16))
        self.bilTypeOutput.setObjectName("bilTypOutput")
        self.numberOfPositionsOutput = QtWidgets.QLabel(bilCard)
        self.numberOfPositionsOutput.setGeometry(QtCore.QRect(400, 100, 47, 16))
        self.numberOfPositionsOutput.setObjectName("numberOfPositionsOutput")
        self.billData=billData
        self.retranslateUi(bilCard)
        QtCore.QMetaObject.connectSlotsByName(bilCard)

    def retranslateUi(self, bilCard):
        _translate = QtCore.QCoreApplication.translate
        bilCard.setWindowTitle(_translate("bilCard", "Счет"))
        self.headerLabel.setText(_translate("bilCard", "Карточка счета "))
        self.companyNameLabel.setText(_translate("bilCard", "Название компании: "))
        self.priceLabel.setText(_translate("bilCard", "Сумма: "))
        self.documentTypeLabel.setText(_translate("bilCard", "Тип документа:"))
        self.documentNumberLabel.setText(_translate("bilCard", "Номер документа: "))
        self.numberOfPositionsLabel.setText(_translate("bilCard", "Количество позиций: "))
        self.bilTypeLabel.setText(_translate("bilCard", "Тип счета: "))
        self.bilTypeOutput.setText(_translate("bilCard", 'text'))
        self.companyNameOutput.setText(self.billData[1])
        self.priceOutput.setText(self.billData[2])
        self.numberOfPositionsOutput.setText(self.billData[5])
        self.bilTypeOutput.setText(self.billData[6])
        if not self.billData[4]:
            self.deliveryYesNoOutput.setText('НЕ ОТГРУЖЕН')
        else:
            self.deliveryYesNoOutput.setText('ОТГРУЖЕН')
            self.documentTypeOutput.setText(self.billData[3])
            self.documentNumberOutput.setText(self.billData[4])