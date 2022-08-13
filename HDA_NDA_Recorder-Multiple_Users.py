# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'HDA_NDA_Recorder_v1.0.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

from PyQt5.QtWidgets import QMessageBox
import os,time,openpyxl,pprint
from decimal import Decimal
App_data_path=os.path.dirname(__file__)+"\\App_data\\"
temp_data_path=os.path.dirname(__file__)+"\\temp_data\\"
generated_files_path=os.path.dirname(__file__)+"\\HDA_NDA_generated_files\\"
monthdict={1:'JAN',2:'FEB',3:'MARCH',4:'APRIL',5:'MAY',6:'JUNE',7:'JULY',8:'AUG',9:'SEPT',10:'OCT',11:'NOV',12:'DEC'}
TO_login=0
#print(TO_login)

class Ui_MainWindow(QtWidgets.QMainWindow):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(823, 594)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 821, 561))
        self.tabWidget.setObjectName("tabWidget")
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.label_26 = QtWidgets.QLabel(self.tab_4)
        self.label_26.setGeometry(QtCore.QRect(160, 80, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_26.setFont(font)
        self.label_26.setObjectName("label_26")
        self.label_27 = QtWidgets.QLabel(self.tab_4)
        self.label_27.setGeometry(QtCore.QRect(590, 90, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_27.setFont(font)
        self.label_27.setObjectName("label_27")
        self.label_28 = QtWidgets.QLabel(self.tab_4)
        self.label_28.setGeometry(QtCore.QRect(100, 150, 47, 13))
        self.label_28.setObjectName("label_28")
        self.label_29 = QtWidgets.QLabel(self.tab_4)
        self.label_29.setGeometry(QtCore.QRect(100, 190, 47, 13))
        self.label_29.setObjectName("label_29")
        self.label_30 = QtWidgets.QLabel(self.tab_4)
        self.label_30.setGeometry(QtCore.QRect(100, 270, 91, 16))
        self.label_30.setObjectName("label_30")
        self.label_31 = QtWidgets.QLabel(self.tab_4)
        self.label_31.setGeometry(QtCore.QRect(510, 190, 91, 16))
        self.label_31.setObjectName("label_31")
        self.label_32 = QtWidgets.QLabel(self.tab_4)
        self.label_32.setGeometry(QtCore.QRect(510, 150, 47, 13))
        self.label_32.setObjectName("label_32")
        self.pushButton_register = QtWidgets.QPushButton(self.tab_4)
        self.pushButton_register.setGeometry(QtCore.QRect(170, 320, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_register.setFont(font)
        self.pushButton_register.setObjectName("pushButton_register")
        self.pushButton_login = QtWidgets.QPushButton(self.tab_4)
        self.pushButton_login.setGeometry(QtCore.QRect(570, 320, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_login.setFont(font)
        self.pushButton_login.setObjectName("pushButton_login")
        self.line = QtWidgets.QFrame(self.tab_4)
        self.line.setGeometry(QtCore.QRect(400, 130, 20, 221))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lineEdit_TOname_step1 = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit_TOname_step1.setGeometry(QtCore.QRect(210, 150, 113, 20))
        self.lineEdit_TOname_step1.setObjectName("lineEdit_TOname_step1")
        self.lineEdit_empNoReg = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit_empNoReg.setGeometry(QtCore.QRect(210, 190, 113, 20))
        self.lineEdit_empNoReg.setObjectName("lineEdit_empNoReg")
        self.lineEdit_empNoLogin = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit_empNoLogin.setGeometry(QtCore.QRect(630, 150, 81, 20))
        self.lineEdit_empNoLogin.setObjectName("lineEdit_empNoLogin")
        self.label_65 = QtWidgets.QLabel(self.tab_4)
        self.label_65.setGeometry(QtCore.QRect(400, 100, 47, 13))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label_65.setFont(font)
        self.label_65.setObjectName("label_65")
        self.label_66 = QtWidgets.QLabel(self.tab_4)
        self.label_66.setGeometry(QtCore.QRect(100, 230, 91, 16))
        self.label_66.setObjectName("label_66")
        self.lineEdit_designation_step1 = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit_designation_step1.setGeometry(QtCore.QRect(210, 230, 111, 20))
        self.lineEdit_designation_step1.setObjectName("lineEdit_designation_step1")
        self.dateEdit_register = QtWidgets.QDateEdit(self.tab_4)
        self.dateEdit_register.setGeometry(QtCore.QRect(210, 270, 81, 22))
        self.dateEdit_register.setCalendarPopup(True)
        self.dateEdit_register.setDate(QtCore.QDate(2019, 1, 1))
        self.dateEdit_register.setObjectName("dateEdit_register")
        self.dateEdit_login = QtWidgets.QDateEdit(self.tab_4)
        self.dateEdit_login.setGeometry(QtCore.QRect(630, 190, 81, 22))
        self.dateEdit_login.setCalendarPopup(True)
        self.dateEdit_login.setDate(QtCore.QDate(2019, 1, 1))
        self.dateEdit_login.setObjectName("dateEdit_login")
        self.tabWidget.addTab(self.tab_4, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.label_9 = QtWidgets.QLabel(self.tab_2)
        self.label_9.setGeometry(QtCore.QRect(290, 310, 71, 21))
        self.label_9.setObjectName("label_9")
        self.pushButton_updateNDA = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_updateNDA.setGeometry(QtCore.QRect(510, 430, 71, 23))
        self.pushButton_updateNDA.setObjectName("pushButton_updateNDA")
        self.label_3 = QtWidgets.QLabel(self.tab_2)
        self.label_3.setGeometry(QtCore.QRect(290, 80, 61, 21))
        self.label_3.setObjectName("label_3")
        self.lineEdit_date = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_date.setGeometry(QtCore.QRect(410, 190, 81, 20))
        self.lineEdit_date.setObjectName("lineEdit_date")
        self.label_7 = QtWidgets.QLabel(self.tab_2)
        self.label_7.setGeometry(QtCore.QRect(290, 250, 51, 21))
        self.label_7.setObjectName("label_7")
        self.lineEdit_remark = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_remark.setGeometry(QtCore.QRect(410, 340, 181, 20))
        self.lineEdit_remark.setObjectName("lineEdit_remark")
        self.label_17 = QtWidgets.QLabel(self.tab_2)
        self.label_17.setGeometry(QtCore.QRect(600, 430, 111, 21))
        self.label_17.setObjectName("label_17")
        self.lineEdit_dutyNo_save = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_dutyNo_save.setGeometry(QtCore.QRect(410, 250, 81, 20))
        self.lineEdit_dutyNo_save.setObjectName("lineEdit_dutyNo_save")
        self.dateEdit_step2 = QtWidgets.QDateEdit(self.tab_2)
        self.dateEdit_step2.setGeometry(QtCore.QRect(410, 20, 81, 22))
        self.dateEdit_step2.setCalendarPopup(True)
        self.dateEdit_step2.setDate(QtCore.QDate(2019, 1, 1))
        self.dateEdit_step2.setObjectName("dateEdit_step2")
        self.label_6 = QtWidgets.QLabel(self.tab_2)
        self.label_6.setGeometry(QtCore.QRect(290, 220, 81, 21))
        self.label_6.setObjectName("label_6")
        self.lineEdit_km = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_km.setGeometry(QtCore.QRect(410, 280, 81, 20))
        self.lineEdit_km.setObjectName("lineEdit_km")
        self.label_14 = QtWidgets.QLabel(self.tab_2)
        self.label_14.setGeometry(QtCore.QRect(640, 400, 71, 21))
        self.label_14.setObjectName("label_14")
        self.label = QtWidgets.QLabel(self.tab_2)
        self.label.setGeometry(QtCore.QRect(290, 190, 41, 21))
        self.label.setObjectName("label")
        self.label_16 = QtWidgets.QLabel(self.tab_2)
        self.label_16.setGeometry(QtCore.QRect(290, 430, 111, 21))
        self.label_16.setObjectName("label_16")
        self.label_12 = QtWidgets.QLabel(self.tab_2)
        self.label_12.setGeometry(QtCore.QRect(640, 370, 71, 21))
        self.label_12.setObjectName("label_12")
        self.comboBox_dutyType = QtWidgets.QComboBox(self.tab_2)
        self.comboBox_dutyType.setGeometry(QtCore.QRect(410, 50, 81, 22))
        self.comboBox_dutyType.setObjectName("comboBox_dutyType")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.comboBox_dutyType.addItem("")
        self.SignOnOffTime_actual = QtWidgets.QTimeEdit(self.tab_2)
        self.SignOnOffTime_actual.setGeometry(QtCore.QRect(410, 400, 81, 22))
        self.SignOnOffTime_actual.setObjectName("SignOnOffTime_actual")
        self.SignOnOffPlace_shed = QtWidgets.QLineEdit(self.tab_2)
        self.SignOnOffPlace_shed.setGeometry(QtCore.QRect(510, 370, 111, 20))
        self.SignOnOffPlace_shed.setObjectName("SignOnOffPlace_shed")
        self.comboBox_tt = QtWidgets.QComboBox(self.tab_2)
        self.comboBox_tt.setGeometry(QtCore.QRect(410, 110, 81, 22))
        self.comboBox_tt.setObjectName("comboBox_tt")
        self.comboBox_tt.addItem("")
        self.comboBox_tt.addItem("")
        self.comboBox_tt.addItem("")
        self.label_10 = QtWidgets.QLabel(self.tab_2)
        self.label_10.setGeometry(QtCore.QRect(290, 340, 71, 21))
        self.label_10.setObjectName("label_10")
        self.pushButton_fetchData = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_fetchData.setGeometry(QtCore.QRect(330, 140, 111, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_fetchData.setFont(font)
        self.pushButton_fetchData.setObjectName("pushButton_fetchData")
        self.lineEdit_dutyNo_fetch = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_dutyNo_fetch.setGeometry(QtCore.QRect(410, 80, 81, 20))
        self.lineEdit_dutyNo_fetch.setObjectName("lineEdit_dutyNo_fetch")
        self.label_5 = QtWidgets.QLabel(self.tab_2)
        self.label_5.setGeometry(QtCore.QRect(290, 20, 41, 21))
        self.label_5.setObjectName("label_5")
        self.lineEdit_dutyType = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_dutyType.setGeometry(QtCore.QRect(410, 220, 81, 20))
        self.lineEdit_dutyType.setObjectName("lineEdit_dutyType")
        self.lineEdit_deviation = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_deviation.setGeometry(QtCore.QRect(410, 310, 81, 20))
        self.lineEdit_deviation.setObjectName("lineEdit_deviation")
        self.label_8 = QtWidgets.QLabel(self.tab_2)
        self.label_8.setGeometry(QtCore.QRect(290, 280, 41, 21))
        self.label_8.setObjectName("label_8")
        self.pushButton_save = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_save.setGeometry(QtCore.QRect(330, 490, 111, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_save.setFont(font)
        self.pushButton_save.setObjectName("pushButton_save")
        self.SignOnOffPlace_actual = QtWidgets.QLineEdit(self.tab_2)
        self.SignOnOffPlace_actual.setGeometry(QtCore.QRect(510, 400, 113, 20))
        self.SignOnOffPlace_actual.setObjectName("SignOnOffPlace_actual")
        self.label_2 = QtWidgets.QLabel(self.tab_2)
        self.label_2.setGeometry(QtCore.QRect(290, 50, 61, 21))
        self.label_2.setObjectName("label_2")
        self.label_15 = QtWidgets.QLabel(self.tab_2)
        self.label_15.setGeometry(QtCore.QRect(200, 390, 81, 21))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_15.setFont(font)
        self.label_15.setObjectName("label_15")
        self.label_11 = QtWidgets.QLabel(self.tab_2)
        self.label_11.setGeometry(QtCore.QRect(290, 370, 91, 21))
        self.label_11.setObjectName("label_11")
        self.label_13 = QtWidgets.QLabel(self.tab_2)
        self.label_13.setGeometry(QtCore.QRect(290, 400, 101, 21))
        self.label_13.setObjectName("label_13")
        self.comboBox_NDA = QtWidgets.QComboBox(self.tab_2)
        self.comboBox_NDA.setGeometry(QtCore.QRect(410, 430, 81, 22))
        self.comboBox_NDA.setObjectName("comboBox_NDA")
        self.comboBox_NDA.addItem("")
        self.comboBox_NDA.addItem("")
        self.comboBox_NDA.addItem("")
        self.comboBox_NDA.addItem("")
        self.comboBox_NDA.addItem("")
        self.SignOnOffTime_shed = QtWidgets.QTimeEdit(self.tab_2)
        self.SignOnOffTime_shed.setGeometry(QtCore.QRect(410, 370, 81, 22))
        self.SignOnOffTime_shed.setCalendarPopup(False)
        self.SignOnOffTime_shed.setTimeSpec(QtCore.Qt.LocalTime)
        self.SignOnOffTime_shed.setObjectName("SignOnOffTime_shed")
        self.label_4 = QtWidgets.QLabel(self.tab_2)
        self.label_4.setGeometry(QtCore.QRect(290, 110, 51, 21))
        self.label_4.setObjectName("label_4")
        self.label_TOnameAlert = QtWidgets.QLabel(self.tab_2)
        self.label_TOnameAlert.setGeometry(QtCore.QRect(10, 20, 221, 51))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_TOnameAlert.setFont(font)
        self.label_TOnameAlert.setWordWrap(True)
        self.label_TOnameAlert.setObjectName("label_TOnameAlert")
        self.label_25 = QtWidgets.QLabel(self.tab_2)
        self.label_25.setGeometry(QtCore.QRect(450, 500, 231, 21))
        self.label_25.setObjectName("label_25")
        self.label_23 = QtWidgets.QLabel(self.tab_2)
        self.label_23.setGeometry(QtCore.QRect(520, 80, 211, 21))
        self.label_23.setObjectName("label_23")
        self.label_24 = QtWidgets.QLabel(self.tab_2)
        self.label_24.setGeometry(QtCore.QRect(520, 110, 161, 21))
        self.label_24.setObjectName("label_24")
        self.label_67 = QtWidgets.QLabel(self.tab_2)
        self.label_67.setGeometry(QtCore.QRect(10, 120, 211, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_67.setFont(font)
        self.label_67.setWordWrap(True)
        self.label_67.setObjectName("label_67")
        self.dateEdit_newMonth = QtWidgets.QDateEdit(self.tab_2)
        self.dateEdit_newMonth.setGeometry(QtCore.QRect(60, 170, 81, 22))
        self.dateEdit_newMonth.setCalendarPopup(True)
        self.dateEdit_newMonth.setDate(QtCore.QDate(2019, 1, 1))
        self.dateEdit_newMonth.setObjectName("dateEdit_newMonth")
        self.pushButton_newMonth = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_newMonth.setGeometry(QtCore.QRect(70, 210, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_newMonth.setFont(font)
        self.pushButton_newMonth.setObjectName("pushButton_newMonth")
        self.line_3 = QtWidgets.QFrame(self.tab_2)
        self.line_3.setGeometry(QtCore.QRect(240, 10, 16, 241))
        self.line_3.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.pushButton_logout = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_logout.setGeometry(QtCore.QRect(720, 490, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_logout.setFont(font)
        self.pushButton_logout.setObjectName("pushButton_logout")
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.pushButton_generateExcel_logout = QtWidgets.QPushButton(self.tab_3)
        self.pushButton_generateExcel_logout.setGeometry(QtCore.QRect(270, 220, 191, 41))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_generateExcel_logout.setFont(font)
        self.pushButton_generateExcel_logout.setObjectName("pushButton_generateExcel_logout")
        self.label_fileLocation = QtWidgets.QLabel(self.tab_3)
        self.label_fileLocation.setGeometry(QtCore.QRect(110, 270, 631, 101))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_fileLocation.setFont(font)
        self.label_fileLocation.setWordWrap(True)
        self.label_fileLocation.setObjectName("label_fileLocation")
        self.label_22 = QtWidgets.QLabel(self.tab_3)
        self.label_22.setGeometry(QtCore.QRect(560, 500, 241, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_22.setFont(font)
        self.label_22.setObjectName("label_22")
        self.label_18 = QtWidgets.QLabel(self.tab_3)
        self.label_18.setGeometry(QtCore.QRect(290, 50, 51, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_18.setFont(font)
        self.label_18.setObjectName("label_18")
        self.label_19 = QtWidgets.QLabel(self.tab_3)
        self.label_19.setGeometry(QtCore.QRect(290, 80, 61, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_19.setFont(font)
        self.label_19.setObjectName("label_19")
        self.label_20 = QtWidgets.QLabel(self.tab_3)
        self.label_20.setGeometry(QtCore.QRect(290, 110, 81, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_20.setFont(font)
        self.label_20.setObjectName("label_20")
        self.label_name_step3 = QtWidgets.QLabel(self.tab_3)
        self.label_name_step3.setGeometry(QtCore.QRect(420, 50, 241, 21))
        self.label_name_step3.setObjectName("label_name_step3")
        self.label_empNo_step3 = QtWidgets.QLabel(self.tab_3)
        self.label_empNo_step3.setGeometry(QtCore.QRect(420, 80, 131, 21))
        self.label_empNo_step3.setObjectName("label_empNo_step3")
        self.label_design_step3 = QtWidgets.QLabel(self.tab_3)
        self.label_design_step3.setGeometry(QtCore.QRect(420, 110, 111, 21))
        self.label_design_step3.setObjectName("label_design_step3")
        self.label_alert_step3 = QtWidgets.QLabel(self.tab_3)
        self.label_alert_step3.setGeometry(QtCore.QRect(230, 140, 321, 61))
        self.label_alert_step3.setObjectName("label_alert_step3")
        self.tabWidget.addTab(self.tab_3, "")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 823, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        #connecting buttons
        self.pushButton_register.clicked.connect(self.register)
        self.pushButton_login.clicked.connect(self.login)
        self.pushButton_newMonth.clicked.connect(self.newMonthHDA_NDA)
        self.pushButton_fetchData.clicked.connect(self.FetchData)
        self.pushButton_save.clicked.connect(self.SaveData)
        self.pushButton_updateNDA.clicked.connect(self.UpdateNDA)
        self.pushButton_logout.clicked.connect(self.logout)
        self.pushButton_generateExcel_logout.clicked.connect(self.GenerateExcel_logout)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "HDA_NDA_Recorder_v1.3"))
        self.label_26.setText(_translate("MainWindow", "Resister Here"))
        self.label_27.setText(_translate("MainWindow", "Login"))
        self.label_28.setText(_translate("MainWindow", "TO Name"))
        self.label_29.setText(_translate("MainWindow", "Emp no"))
        self.label_30.setText(_translate("MainWindow", "Date of birth"))
        self.label_31.setText(_translate("MainWindow", "Date of birth"))
        self.label_32.setText(_translate("MainWindow", "Emp no"))
        self.pushButton_register.setText(_translate("MainWindow", "Register"))
        self.pushButton_login.setText(_translate("MainWindow", "Login"))
        self.label_65.setText(_translate("MainWindow", "Or"))
        self.label_66.setText(_translate("MainWindow", "Designation"))
        self.dateEdit_register.setDisplayFormat(_translate("MainWindow", "d/M/yyyy"))
        self.dateEdit_login.setDisplayFormat(_translate("MainWindow", "d/M/yyyy"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_4), _translate("MainWindow", "Step 1"))
        self.label_9.setText(_translate("MainWindow", "Deviation"))
        self.pushButton_updateNDA.setText(_translate("MainWindow", "Update NDA "))
        self.label_3.setText(_translate("MainWindow", "Duty no"))
        self.label_7.setText(_translate("MainWindow", "Duty no"))
        self.label_17.setText(_translate("MainWindow", "If actual time changes"))
        self.dateEdit_step2.setDisplayFormat(_translate("MainWindow", "d/M/yyyy"))
        self.label_6.setText(_translate("MainWindow", "Duty Type"))
        self.label_14.setText(_translate("MainWindow", "Place"))
        self.label.setText(_translate("MainWindow", "Date"))
        self.label_16.setText(_translate("MainWindow", "NDA"))
        self.label_12.setText(_translate("MainWindow", "Place"))
        self.comboBox_dutyType.setItemText(0, _translate("MainWindow", "RUN"))
        self.comboBox_dutyType.setItemText(1, _translate("MainWindow", "PRT"))
        self.comboBox_dutyType.setItemText(2, _translate("MainWindow", "WD"))
        self.comboBox_dutyType.setItemText(3, _translate("MainWindow", "NIGHT"))
        self.comboBox_dutyType.setItemText(4, _translate("MainWindow", "REST"))
        self.comboBox_dutyType.setItemText(5, _translate("MainWindow", "CR"))
        self.comboBox_dutyType.setItemText(6, _translate("MainWindow", "CL"))
        self.comboBox_dutyType.setItemText(7, _translate("MainWindow", "EL"))
        self.comboBox_dutyType.setItemText(8, _translate("MainWindow", "GH"))
        self.comboBox_dutyType.setItemText(9, _translate("MainWindow", "RH"))
        self.comboBox_dutyType.setItemText(10, _translate("MainWindow", "SICK"))
        self.comboBox_dutyType.setItemText(11, _translate("MainWindow", "OTHER"))
        self.SignOnOffTime_actual.setDisplayFormat(_translate("MainWindow", "hh:mm"))
        self.comboBox_tt.setItemText(0, _translate("MainWindow", "WEEK"))
        self.comboBox_tt.setItemText(1, _translate("MainWindow", "SAT"))
        self.comboBox_tt.setItemText(2, _translate("MainWindow", "SUN"))
        self.label_10.setText(_translate("MainWindow", "Remark"))
        self.pushButton_fetchData.setText(_translate("MainWindow", "Fetch Data"))
        self.label_5.setText(_translate("MainWindow", "Date"))
        self.label_8.setText(_translate("MainWindow", "Km"))
        self.pushButton_save.setText(_translate("MainWindow", "Save"))
        self.label_2.setText(_translate("MainWindow", "Duty Type"))
        self.label_15.setText(_translate("MainWindow", "To claim NDA"))
        self.label_11.setText(_translate("MainWindow", "Sign On/Off Shed"))
        self.label_13.setText(_translate("MainWindow", "Sign On/Off Actual"))
        self.comboBox_NDA.setItemText(0, _translate("MainWindow", "NA"))
        self.comboBox_NDA.setItemText(1, _translate("MainWindow", "A"))
        self.comboBox_NDA.setItemText(2, _translate("MainWindow", "B"))
        self.comboBox_NDA.setItemText(3, _translate("MainWindow", "C"))
        self.comboBox_NDA.setItemText(4, _translate("MainWindow", "D"))
        self.SignOnOffTime_shed.setDisplayFormat(_translate("MainWindow", "hh:mm"))
        self.label_4.setText(_translate("MainWindow", "Time Table"))
        self.label_TOnameAlert.setText(_translate("MainWindow", "Kindly Login to save your data"))
        self.label_25.setText(_translate("MainWindow", "Move to step 3 if all entries made for the month"))
        self.label_23.setText(_translate("MainWindow", "Leave it blank if not available/required"))
        self.label_24.setText(_translate("MainWindow", "Leave it as it is if not required"))
        self.label_67.setText(_translate("MainWindow", "To record new months data, choose first date of month"))
        self.dateEdit_newMonth.setDisplayFormat(_translate("MainWindow", "d/M/yyyy"))
        self.pushButton_newMonth.setText(_translate("MainWindow", "Submit"))
        self.pushButton_logout.setText(_translate("MainWindow", "Log Out"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Step 2"))
        self.pushButton_generateExcel_logout.setText(_translate("MainWindow", "Generate excel file and Log Out"))
        self.label_fileLocation.setText(_translate("MainWindow", "?"))
        self.label_22.setText(_translate("MainWindow", "Developer: satishkushwah50@gmail.com"))
        self.label_18.setText(_translate("MainWindow", "Name:"))
        self.label_19.setText(_translate("MainWindow", "Emp no:"))
        self.label_20.setText(_translate("MainWindow", "Designation:"))
        self.label_name_step3.setText(_translate("MainWindow", "?"))
        self.label_empNo_step3.setText(_translate("MainWindow", "?"))
        self.label_design_step3.setText(_translate("MainWindow", "?"))
        self.label_alert_step3.setText(_translate("MainWindow", "?"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "Step 3"))

        #loading temp data
        self.label_fileLocation.setText("Generated files will be saved in "+generated_files_path)

    def register(self):
        try:
            from App_data import Registered_TO
            database_dict=Registered_TO.TO_database
            dataList=[]
            dataList.append(self.lineEdit_TOname_step1.text().strip())
            dataList.append(int(self.lineEdit_empNoReg.text().strip()))
            dataList.append(self.lineEdit_designation_step1.text().strip())
            dataList.append(self.dateEdit_register.text())
            res = QMessageBox.question(self,'Registration Data','Register with following data?\n\n'+'Name:                 '+dataList[0]+'\nEmp no:              '+str(dataList[1])+\
                '\nDesignation:       '+dataList[2]+'\nDate of Birth:      '+dataList[3], QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if res==QMessageBox.Yes:
                database_dict[dataList[1]]=dataList
                fp=open(App_data_path+'Registered_TO.py','w'); fp.write('TO_database'+'='+pprint.pformat(database_dict)); fp.close()
                hdafile=open(temp_data_path+str(dataList[1])+'_HDA.csv','w'); hdafile.write('Date,Duty no,Duty Type,Km,Deviation,Actual Km\n'); hdafile.close()
                ndafile=open(temp_data_path+str(dataList[1])+'_NDA.csv','w'); ndafile.write('Date,Duty no,SignOn/Off Time Schuled,SignOn/Off Place Scheuled,SignOn/Off Time Actual,SignOn/Off Place Actual,Full Night,Before 0400 or after 0000,0401-0500 or 2301-2359,0501-0600 or 2201-2300,Remarks\n'); ndafile.close()  
                QMessageBox.about(self,'Registration Successful',dataList[0]+", You are registered Successfully")
                self.lineEdit_TOname_step1.setText('')
                self.lineEdit_designation_step1.setText('')
                self.lineEdit_empNoReg.setText('')
                self.dateEdit_register.setDate(QtCore.QDate(2019,1,1))
        except:
            QMessageBox.about(self,'Error','Some error occured')
    def login(self):
        from App_data import Registered_TO
        database_dict=Registered_TO.TO_database
        try:
            dataList=database_dict[int(self.lineEdit_empNoLogin.text().strip())]
            if dataList[1] in database_dict.keys():
                if dataList[3]==self.dateEdit_login.text():
                    self.lineEdit_empNoLogin.setText('')
                    self.dateEdit_login.setDate(QtCore.QDate(2019,1,1))
                    global TO_login
                    TO_login=dataList[1] 
                    #print(TO_login) 
                    fp=open(temp_data_path+str(dataList[1])+'_HDA.csv','r'); l=len(fp.readlines()); fp.close()  
                    self.label_name_step3.setText(dataList[0])
                    self.label_empNo_step3.setText(str(dataList[1]))
                    self.label_design_step3.setText(dataList[2])
                    self.label_alert_step3.setText(dataList[0]+', you have made '+str(l-1)+' entries')
                    
                    if l==1:
                        self.label_TOnameAlert.setText(dataList[0]+', kindly choose month to record HDA NDA data')
                    elif l>1:
                        fp=open(temp_data_path+str(dataList[1])+'_HDA.csv','r'); hdadata=fp.readlines(); l=len(hdadata)
                        #print(hdadata)
                        self.label_TOnameAlert.setText(dataList[0]+', you have made '+str(l-1)+' entries in '+monthdict[int(hdadata[1].split(',')[0].split('/')[1])]+' '+hdadata[1].split(',')[0].split('/')[2])
                        self.dateEdit_step2.setDate(QtCore.QDate(int(hdadata[1].split(',')[0].split('/')[2]),int(hdadata[1].split(',')[0].split('/')[1]),l))
                        fp.close()
                    QMessageBox.about(self,'Login Successful',"Welcome "+dataList[0]+', you are ready to record your HDA NDA data, move to step 2')
                else:
                    QMessageBox.about(self,'Login Failed',"Date of birth not matched")
            else:
                QMessageBox.about(self,'Login Failed',"You are not registered ")
        except KeyError:
            QMessageBox.about(self,'Login Failed',"You are not registered or entering wrong combination of data")

    def newMonthHDA_NDA(self):
        #print(TO_login)
        if TO_login!=0:
            self.label_TOnameAlert.setText(self.label_name_step3.text()+', data being recorded of '+monthdict[int(self.dateEdit_newMonth.text().split('/')[1])]+' '+self.dateEdit_newMonth.text().split('/')[2]+', 0 entries made')
            self.dateEdit_step2.setDate(QtCore.QDate(int(self.dateEdit_newMonth.text().split('/')[2]),int(self.dateEdit_newMonth.text().split('/')[1]),1))
            hdafile=open(temp_data_path+str(TO_login)+'_HDA.csv','w'); hdafile.write('Date,Duty no,Duty Type,Km,Deviation,Actual Km\n'); hdafile.close()
            ndafile=open(temp_data_path+str(TO_login)+'_NDA.csv','w'); ndafile.write('Date,Duty no,SignOn/Off Time Schuled,SignOn/Off Place Scheuled,SignOn/Off Time Actual,SignOn/Off Place Actual,Full Night,Before 0400 or after 0000,0401-0500 or 2301-2359,0501-0600 or 2201-2300,Remarks\n'); ndafile.close()  
            self.label_alert_step3.setText('?')
    def FetchData(self):
        dutyno=self.lineEdit_dutyNo_fetch.text().strip()
        if dutyno!='':
            km_file = openpyxl.load_workbook(App_data_path+'Duty_km_data.xlsx')
            signonoff_file=openpyxl.load_workbook(App_data_path+'Duty_SignOnOff_data.xlsx')
            if self.comboBox_tt.currentText()=='WEEK':
                kmsheet=km_file['WEEK']
                signonoffsheet=signonoff_file['WEEK']  
            elif self.comboBox_tt.currentText()=='SAT':
                kmsheet=km_file['SAT']
                signonoffsheet=signonoff_file['SAT']
            elif self.comboBox_tt.currentText()=='SUN':
                kmsheet=km_file['SUN']
                signonoffsheet=signonoff_file['SUN']

            self.lineEdit_date.setText(self.dateEdit_step2.text())
            self.lineEdit_dutyType.setText(self.comboBox_dutyType.currentText())
            self.lineEdit_dutyNo_save.setText(dutyno)
            self.lineEdit_km.setText(str(kmsheet.cell(row=int(dutyno)+1,column=2).value))
            self.lineEdit_deviation.setText('0')
            self.lineEdit_remark.setText('-')
            
            SignOn=str(signonoffsheet.cell(column=2,row=int(dutyno)+1).value)[-8:-3]; SignOff=str(signonoffsheet.cell(column=4,row=int(dutyno)+1).value)[-8:-3]
            if len(SignOn)==4: SignOn='0'+SignOn
            if len(SignOff)==4: SignOff='0'+SignOff
            if '04:00'<=SignOn<='06:00':
                self.SignOnOffTime_shed.setTime(QtCore.QTime(int(SignOn.split(':')[0]),int(SignOn.split(':')[1]),00))
                self.SignOnOffPlace_shed.setText(str(signonoffsheet.cell(column=3,row=int(dutyno)+1).value))
                self.SignOnOffTime_actual.setTime(QtCore.QTime(int(SignOn.split(':')[0]),int(SignOn.split(':')[1]),00))
                self.SignOnOffPlace_actual.setText(str(signonoffsheet.cell(column=3,row=int(dutyno)+1).value))     
            else:
                self.SignOnOffTime_shed.setTime(QtCore.QTime(int(SignOff.split(':')[0]),int(SignOff.split(':')[1]),00))
                self.SignOnOffPlace_shed.setText(str(signonoffsheet.cell(column=5,row=int(dutyno)+1).value))
                self.SignOnOffTime_actual.setTime(QtCore.QTime(int(SignOff.split(':')[0]),int(SignOff.split(':')[1]),00))
                self.SignOnOffPlace_actual.setText(str(signonoffsheet.cell(column=5,row=int(dutyno)+1).value))
           
            SignOnOffTimeAct=self.SignOnOffTime_shed.text() #to calculate NDA
            if self.lineEdit_dutyType.text()=='NIGHT': 
                self.comboBox_NDA.setCurrentIndex(1)
            elif '00:00'<=SignOnOffTimeAct<='02:00': 
                self.comboBox_NDA.setCurrentIndex(2)
            elif '04:00'<SignOnOffTimeAct<='05:00' or '23:01'<=SignOnOffTimeAct<='23:59': 
                self.comboBox_NDA.setCurrentIndex(3)
            elif '05:01'<=SignOnOffTimeAct<='06:00' or '22:01'<=SignOnOffTimeAct<='23:00': 
                self.comboBox_NDA.setCurrentIndex(4)
            else: 
                self.comboBox_NDA.setCurrentIndex(0)  
        else:
            self.lineEdit_date.setText(self.dateEdit_step2.text())
            self.lineEdit_dutyType.setText(self.comboBox_dutyType.currentText())
            self.lineEdit_dutyNo_save.setText('-')
            self.lineEdit_km.setText('0')
            self.lineEdit_deviation.setText('0')
            self.lineEdit_remark.setText('-')
            self.SignOnOffTime_shed.setTime(QtCore.QTime(00,00,00))
            self.SignOnOffPlace_shed.setText('NA')
            self.SignOnOffTime_actual.setTime(QtCore.QTime(00,00,00))
            self.SignOnOffPlace_actual.setText('NA')
            if self.lineEdit_dutyType.text()=='NIGHT': self.comboBox_NDA.setCurrentIndex(1)
            else:   self.comboBox_NDA.setCurrentIndex(0)

    def UpdateNDA(self):
        SignOnOffTimeAct=self.SignOnOffTime_actual.text()
        if self.lineEdit_dutyType.text()=='NIGHT': 
            self.comboBox_NDA.setCurrentIndex(1)
        elif '00:00'<=SignOnOffTimeAct<='02:00': 
            self.comboBox_NDA.setCurrentIndex(2)
        elif '04:00'<SignOnOffTimeAct<='05:00' or '23:01'<=SignOnOffTimeAct<='23:59': 
            self.comboBox_NDA.setCurrentIndex(3)
        elif '05:01'<=SignOnOffTimeAct<='06:00' or '22:01'<=SignOnOffTimeAct<='23:00': 
            self.comboBox_NDA.setCurrentIndex(4)
        else: 
            self.comboBox_NDA.setCurrentIndex(0)
    
    def SaveData(self):
        if TO_login!=0:
            dutydate=self.lineEdit_date.text()
            dutyno=self.lineEdit_dutyNo_save.text()
            dutytype=self.lineEdit_dutyType.text()
            dutykm=self.lineEdit_km.text()
            deviation=self.lineEdit_deviation.text()
            remark=self.lineEdit_remark.text()
            SignOnOff_Shed=self.SignOnOffTime_shed.text()+'/'+self.SignOnOffPlace_shed.text()
            SignOnOff_Actual=self.SignOnOffTime_actual.text()+'/'+self.SignOnOffPlace_actual.text()
            if dutyno=='-' and dutytype not in ['RUN','PRT','WD','NIGHT']:
                SignOnOff_Shed='NA/NA'; SignOnOff_Actual='NA/NA'

            nda=self.comboBox_NDA.currentText()
            nda1='-'; nda2='-'; nda3='-'; nda4='-'
            if nda=='A': nda1='Y'    
            elif nda=='B': nda2='Y'
            elif nda=='C': nda3='Y'
            elif nda=='D': nda4='Y'
            else: 
                ndastr='-,-,-,-,-,-,-,-,-,-,'+remark+'\n'
            
            res = QMessageBox.question(self,'Duty Data','Save following data?\n\n'+'Date:                 '+dutydate+'\nDutyNo:            '+dutyno+\
                '\nDutyType:        '+dutytype+'\nKm:                   '+dutykm+'\nDeviation:         '+deviation+'\nActualKm:        '+str(round(Decimal(dutykm)+Decimal(deviation),1))+\
                '\nRemark:             '+remark+'\n'+'SignOn/Off_Shed:  '+SignOnOff_Shed+'\n'+'SignOn/Off_Actual:  '+SignOnOff_Actual+\
                '\nNDA:                  '+nda, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if res==QMessageBox.Yes:
                file=open(temp_data_path+str(TO_login)+'_HDA.csv','a')
                file.write(dutydate+','+dutyno+','+dutytype+','+dutykm+','+deviation+','+str(round(Decimal(dutykm)+Decimal(deviation),1))+'\n')
                file.close()
                file=open(temp_data_path+str(TO_login)+'_NDA.csv','a')
                if nda=='NA':
                    file.write(ndastr)
                else:
                    file.write(dutydate+','+dutyno+','+self.SignOnOffTime_shed.text()+','+self.SignOnOffPlace_shed.text()+\
                    ','+self.SignOnOffTime_actual.text()+','+self.SignOnOffPlace_actual.text()+','+nda1+','+nda2+','+nda3+','+nda4+','+remark+'\n')
                file.close()
                QMessageBox.about(self,'Data Saved','Data saved successfully')
                file=open(temp_data_path+str(TO_login)+'_HDA.csv','r'); hdadata=file.readlines(); l=len(hdadata); file.close(); 
                self.dateEdit_step2.setDate(QtCore.QDate(int(self.dateEdit_step2.text().split('/')[2]), int(self.dateEdit_step2.text().split('/')[1]), l))
                self.lineEdit_date.setText(self.dateEdit_step2.text())
                self.label_TOnameAlert.setText(self.label_name_step3.text()+', you have made '+str(l-1)+' entries in '+monthdict[int(hdadata[1].split(',')[0].split('/')[1])]+' '+hdadata[1].split(',')[0].split('/')[2])
                self.label_alert_step3.setText(self.label_name_step3.text()+', you have made '+str(l-1)+' entries in '+monthdict[int(hdadata[1].split(',')[0].split('/')[1])]+' '+hdadata[1].split(',')[0].split('/')[2])
            self.lineEdit_dutyNo_fetch.setText('')

        else:
            QMessageBox.about(self,'Error','Kindly login to save your data')

    def logout(self):
        global TO_login
        TO_login=0
        self.label_name_step3.setText('?')
        self.label_empNo_step3.setText('?')
        self.label_design_step3.setText('?')
        self.label_alert_step3.setText('?')
        self.label_TOnameAlert.setText("Kindly Login to save your data")

    def GenerateExcel_logout(self):
        global TO_login
        if TO_login!=0:

            hdanda = openpyxl.load_workbook(App_data_path+'HDA_NDA_Format.xlsx')
            f=open(temp_data_path+str(TO_login)+'_HDA.csv','r'); f1=open(temp_data_path+str(TO_login)+'_NDA.csv','r')
            hdadata=f.readlines(); ndadata=f1.readlines()
            hdasheet = hdanda['HDA']; ndasheet=hdanda['NDA']
            totalkm=0; runduty=0; fixedsignon=0; A=0; B=0; C=0; D=0
            ndasheet['I1']=hdasheet['G1']='MONTH: '+monthdict[int(self.dateEdit_step2.text().split('/')[1])]+' '+self.dateEdit_step2.text().split('/')[2]
            ndasheet['A2']=hdasheet['A2']='NAME: '+self.label_name_step3.text()
            ndasheet['E2']=hdasheet['C2']='EMP NO: '+self.label_empNo_step3.text()
            ndasheet['G2']=hdasheet['E2']='DESG: '+self.label_design_step3.text()
            for i in range(len(hdadata)-1):
                for j in range(6):
                    hdasheet.cell(row=13+i,column=j+1).value=hdadata[i+1].split(',')[j]
                totalkm+=float(hdadata[i+1].split(',')[5])
                if hdadata[i+1].split(',')[2]=='RUN':
                    runduty+=1
                if hdadata[i+1].split(',')[2] in ['WD', 'PRT', 'NIGHT']:
                    fixedsignon+=1
                for j in range(11):
                    ndasheet.cell(row=13+i,column=j+1).value=ndadata[i+1].split(',')[j]
                if ndadata[i+1].split(',')[6]=='Y':
                    A+=1
                if ndadata[i+1].split(',')[7]=='Y':
                    B+=1
                if ndadata[i+1].split(',')[8]=='Y':
                    C+=1
                if ndadata[i+1].split(',')[9]=='Y':
                    D+=1
            hdasheet['F44']=totalkm; hdasheet['F45']=runduty; hdasheet['F46']=fixedsignon; hdasheet['F47']=runduty+fixedsignon
            ndasheet['G44']=A; ndasheet['H44']=B; ndasheet['I44']=C; ndasheet['J44']=D
            hdanda.save(generated_files_path+self.label_name_step3.text()+'_HDA_NDA_'+monthdict[int(self.dateEdit_step2.text().split('/')[1])]+'_'+self.dateEdit_step2.text().split('/')[2]+".xlsx")
            
            TO_login=0
            self.label_name_step3.setText('?')
            self.label_empNo_step3.setText('?')
            self.label_design_step3.setText('?')
            self.label_alert_step3.setText('?')
            self.label_TOnameAlert.setText("Kindly Login to save your data")
            os.startfile(generated_files_path)
        else:
            QMessageBox.about(self,'Error','Kindly login to creat excel file')
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

