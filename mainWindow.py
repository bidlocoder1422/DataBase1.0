# -*- coding: utf-8 -*-

# добавить скан договора к компании %3 отдел%
# mainWindow : Проверить на Ошибки парсинг xls file %отдел прикладного тестирования%
from PyQt5 import QtCore, QtGui, QtWidgets
import sys,xlrd
from companyAbout import Ui_aboutCompanyWindow
from addcompany import addCompanyDialog
from bilCard import Ui_bilCard
from editCompany import editCompanyDialog


def getCompanyList():
    companiesListFile=xlrd.open_workbook('data/companiesBase.xlsx')
    workSheet=companiesListFile.sheet_by_index(0)
    companies=[]
    for i in range(1,workSheet.nrows):
        companies.append(str(workSheet.cell(i,1).value))
    return companies

def getNumberOfBills():
    BillsListFile = xlrd.open_workbook('data/billsBase.xlsx')
    workSheet = BillsListFile.sheet_by_index(0)
    bills = []
    for i in range(1, workSheet.nrows):
        bills.append('Счет '+str(int(workSheet.cell(i,0 ).value))+' , ' +workSheet.cell(i,1).value)
    return bills

def getBillData(rowNumber):
    BillsListFile = xlrd.open_workbook('data/billsBase.xlsx')
    workSheet = BillsListFile.sheet_by_index(0)
    billData=[]
    for i in range(0,7):
        billData.append(str(workSheet.cell(rowNumber, i).value))
        print(type(workSheet.cell(rowNumber, i).value))
    return billData

class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(640, 290)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.downloadXLSBtn = QtWidgets.QPushButton(self.centralwidget)
        self.downloadXLSBtn.setGeometry(QtCore.QRect(320, 10, 261, 23))
        self.downloadXLSBtn.setObjectName("downloadXLSBtn")
        self.uploadXLSBtn = QtWidgets.QPushButton(self.centralwidget)
        self.uploadXLSBtn.setGeometry(QtCore.QRect(10, 10, 301, 23))
        self.uploadXLSBtn.setObjectName("uploadXLSBtn")
        self.companyListCB = QtWidgets.QComboBox(self.centralwidget)
        self.companyListCB.setGeometry(QtCore.QRect(10, 50, 291, 22))
        self.companyListCB.setObjectName("companyListCB")
        self.openCompanyCardBtn = QtWidgets.QPushButton(self.centralwidget)
        self.openCompanyCardBtn.setGeometry(QtCore.QRect(320, 50, 161, 23))
        self.openCompanyCardBtn.setObjectName("openCompoanyCardBtn")
        self.openBilCarBtn = QtWidgets.QPushButton(self.centralwidget)
        self.openBilCarBtn.setGeometry(QtCore.QRect(320, 90, 161, 23))
        self.openBilCarBtn.setObjectName("openBilCarBtn")
        self.bilChooseCb = QtWidgets.QComboBox(self.centralwidget)
        self.bilChooseCb.setGeometry(QtCore.QRect(10, 90, 291, 22))
        self.bilChooseCb.setObjectName("bilChooseCb")
        self.addCompanyBtn = QtWidgets.QPushButton(self.centralwidget)
        self.addCompanyBtn.setGeometry(QtCore.QRect(10, 130, 591, 23))
        self.addCompanyBtn.setObjectName("addCompanyBtn")
        self.editCompanyBtn = QtWidgets.QPushButton(self.centralwidget)
        self.editCompanyBtn.setGeometry(QtCore.QRect(480, 50, 151, 23))
        self.editCompanyBtn.setObjectName("editCompanyBtn")
        self.UploadExitBtn = QtWidgets.QPushButton(self.centralwidget)
        self.UploadExitBtn.setGeometry(QtCore.QRect(400, 210, 221, 23))
        self.UploadExitBtn.setObjectName("UploadExitBtn")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.menu.addAction(self.actionAbout)
        self.menu.addAction(self.actionExit)
        self.menubar.addAction(self.menu.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Учет"))
        self.downloadXLSBtn.setText(_translate("MainWindow", "Загрузить изменения на компьютер"))
        self.uploadXLSBtn.setText(_translate("MainWindow", "Выгрузить изменения на сервер"))
        self.openCompanyCardBtn.setText(_translate("MainWindow", "Открыть карточку"))
        self.openBilCarBtn.setText(_translate("MainWindow", "Открыть краткую сводку"))
        self.addCompanyBtn.setText(_translate("MainWindow", "Добавить компанию"))
        self.editCompanyBtn.setText(_translate("MainWindow", "Изменить данные"))
        self.UploadExitBtn.setText(_translate("MainWindow", "Выгрузить и выйти"))
        self.menu.setTitle(_translate("MainWindow", "Меню"))
        self.actionAbout.setText(_translate("MainWindow", "About"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.openCompanyCardBtn.clicked.connect(self.getCompanyCard)
        self.addCompanyBtn.clicked.connect(self.addCompany)
        self.openBilCarBtn.clicked.connect(self.getBilCard)
        self.UploadExitBtn.clicked.connect(self.exitUpdateBtnClicked)
        self.editCompanyBtn.clicked.connect(self.editCompanyBtnClicked)
        self.uploadXLSBtn.clicked.connect(self.empttyBtnClicked)
        self.downloadXLSBtn.clicked.connect(self.empttyBtnClicked)

    def updateCompaniesList(self):
        self.companyListCB.clear()
        self.companyList=getCompanyList()
        self.companyListCB.addItems(self.companyList)

    def updateBillsList(self):
        self.bilChooseCb.clear()
        self.bilChooseCb.addItems(getNumberOfBills())

    @QtCore.pyqtSlot()
    def getCompanyCard(self):
        companyCardWindow=QtWidgets.QDialog()
        companyCard=Ui_aboutCompanyWindow()
        companyCard.setupUi(companyCardWindow, self.companyListCB.currentText())
        companyCardWindow.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        companyCard.getInfoFromFile(self.companyListCB.currentText())
        companyCardWindow.exec_()

    @QtCore.pyqtSlot()
    def addCompany(self):
        addCompanyWindow=QtWidgets.QDialog()
        addCompany=addCompanyDialog()
        addCompany.setupUi(addCompanyWindow)
        addCompanyWindow.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        addCompanyWindow.exec_()
        self.updateCompaniesList()

    @QtCore.pyqtSlot()
    def getBilCard(self):
        print(self.bilChooseCb.currentIndex()+1)
        self.billData=getBillData(self.bilChooseCb.currentIndex() + 1)
        print(self.billData)
        bilCardWindow=QtWidgets.QDialog()
        bilCard=Ui_bilCard()
        bilCard.setupUi(bilCardWindow, self.billData)
        bilCardWindow.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        bilCardWindow.exec_()

    @QtCore.pyqtSlot()
    def editCompanyBtnClicked(self):
        editCompanyWindow=QtWidgets.QDialog()
        editCompany=editCompanyDialog()
        editCompany.setupUi(editCompanyWindow)
        editCompanyWindow.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        editCompany.setCompanyData(self.companyListCB.currentText())
        editCompanyWindow.exec_()

    @QtCore.pyqtSlot()
    def exitUpdateBtnClicked(self):
        app.exit()

    @QtCore.pyqtSlot()
    def empttyBtnClicked(self):
        message=QtWidgets.QMessageBox()
        message.setIcon(QtWidgets.QMessageBox.Information)
        message.setText('В следующих версиях')
        message.exec_()

if __name__ == '__main__':
    app=QtWidgets.QApplication(sys.argv)
    MainWindow=QtWidgets.QMainWindow()
    MainWindowObject=Ui_MainWindow()
    MainWindowObject.setupUi(MainWindow)
    MainWindowObject.updateCompaniesList()
    MainWindowObject.updateBillsList()
    MainWindow.show()
    sys.exit(app.exec_())

