# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GUI.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(826, 619)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.tabWidget.setFont(font)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.gridLayout = QtWidgets.QGridLayout(self.tab)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout_15 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_15.setObjectName("horizontalLayout_15")
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox.setFont(font)
        self.groupBox.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox.setFlat(False)
        self.groupBox.setObjectName("groupBox")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.groupBox)
        self.verticalLayout.setObjectName("verticalLayout")
        self.pushButton = QtWidgets.QPushButton(self.groupBox)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.pushButton_3 = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_3.setObjectName("pushButton_3")
        self.verticalLayout.addWidget(self.pushButton_3)
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout.addWidget(self.pushButton_2)
        self.horizontalLayout_15.addWidget(self.groupBox)
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox_2.setFlat(False)
        self.groupBox_2.setObjectName("groupBox_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.groupBox_2)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.pushButton_4 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_4.setObjectName("pushButton_4")
        self.verticalLayout_2.addWidget(self.pushButton_4)
        self.pushButton_6 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_6.setObjectName("pushButton_6")
        self.verticalLayout_2.addWidget(self.pushButton_6)
        self.pushButton_5 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_5.setObjectName("pushButton_5")
        self.verticalLayout_2.addWidget(self.pushButton_5)
        self.pushButton_18 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_18.setObjectName("pushButton_18")
        self.verticalLayout_2.addWidget(self.pushButton_18)
        self.horizontalLayout_15.addWidget(self.groupBox_2)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_4.setFont(font)
        self.groupBox_4.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox_4.setFlat(False)
        self.groupBox_4.setObjectName("groupBox_4")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.groupBox_4)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.pushButton_11 = QtWidgets.QPushButton(self.groupBox_4)
        self.pushButton_11.setObjectName("pushButton_11")
        self.verticalLayout_4.addWidget(self.pushButton_11)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_4 = QtWidgets.QLabel(self.groupBox_4)
        self.label_4.setTextFormat(QtCore.Qt.AutoText)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_2.addWidget(self.label_4)
        self.lineEdit_4 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.horizontalLayout_2.addWidget(self.lineEdit_4)
        self.verticalLayout_4.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_5 = QtWidgets.QLabel(self.groupBox_4)
        self.label_5.setTextFormat(QtCore.Qt.AutoText)
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_3.addWidget(self.label_5)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.horizontalLayout_3.addWidget(self.lineEdit_5)
        self.verticalLayout_4.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_6 = QtWidgets.QLabel(self.groupBox_4)
        self.label_6.setTextFormat(QtCore.Qt.AutoText)
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_4.addWidget(self.label_6)
        self.lineEdit_6 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.horizontalLayout_4.addWidget(self.lineEdit_6)
        self.verticalLayout_4.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.pushButton_12 = QtWidgets.QPushButton(self.groupBox_4)
        self.pushButton_12.setObjectName("pushButton_12")
        self.horizontalLayout_5.addWidget(self.pushButton_12)
        self.pushButton_13 = QtWidgets.QPushButton(self.groupBox_4)
        self.pushButton_13.setObjectName("pushButton_13")
        self.horizontalLayout_5.addWidget(self.pushButton_13)
        self.verticalLayout_4.addLayout(self.horizontalLayout_5)
        self.horizontalLayout_15.addWidget(self.groupBox_4)
        self.gridLayout.addLayout(self.horizontalLayout_15, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout(self.tab_2)
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox_3.setFlat(False)
        self.groupBox_3.setObjectName("groupBox_3")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout(self.groupBox_3)
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_2 = QtWidgets.QLabel(self.groupBox_3)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_3.addWidget(self.label_2)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.verticalLayout_3.addWidget(self.lineEdit_2)
        self.label_3 = QtWidgets.QLabel(self.groupBox_3)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_3.addWidget(self.label_3)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.verticalLayout_3.addWidget(self.lineEdit_3)
        self.pushButton_7 = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_7.setObjectName("pushButton_7")
        self.verticalLayout_3.addWidget(self.pushButton_7)
        self.pushButton_8 = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_8.setObjectName("pushButton_8")
        self.verticalLayout_3.addWidget(self.pushButton_8)
        self.verticalLayout_9.addLayout(self.verticalLayout_3)
        self.horizontalLayout_13.addWidget(self.groupBox_3)
        self.groupBox_5 = QtWidgets.QGroupBox(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_5.setFont(font)
        self.groupBox_5.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox_5.setFlat(False)
        self.groupBox_5.setObjectName("groupBox_5")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout(self.groupBox_5)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.pushButton_16 = QtWidgets.QPushButton(self.groupBox_5)
        self.pushButton_16.setObjectName("pushButton_16")
        self.verticalLayout_5.addWidget(self.pushButton_16)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_7 = QtWidgets.QLabel(self.groupBox_5)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_6.addWidget(self.label_7)
        self.lineEdit_9 = QtWidgets.QLineEdit(self.groupBox_5)
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.horizontalLayout_6.addWidget(self.lineEdit_9)
        self.verticalLayout_5.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_9 = QtWidgets.QLabel(self.groupBox_5)
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_7.addWidget(self.label_9)
        self.lineEdit_10 = QtWidgets.QLineEdit(self.groupBox_5)
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.horizontalLayout_7.addWidget(self.lineEdit_10)
        self.verticalLayout_5.addLayout(self.horizontalLayout_7)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_10 = QtWidgets.QLabel(self.groupBox_5)
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_8.addWidget(self.label_10)
        self.lineEdit_11 = QtWidgets.QLineEdit(self.groupBox_5)
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.horizontalLayout_8.addWidget(self.lineEdit_11)
        self.verticalLayout_5.addLayout(self.horizontalLayout_8)
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.pushButton_14 = QtWidgets.QPushButton(self.groupBox_5)
        self.pushButton_14.setObjectName("pushButton_14")
        self.horizontalLayout_9.addWidget(self.pushButton_14)
        self.pushButton_15 = QtWidgets.QPushButton(self.groupBox_5)
        self.pushButton_15.setObjectName("pushButton_15")
        self.horizontalLayout_9.addWidget(self.pushButton_15)
        self.verticalLayout_5.addLayout(self.horizontalLayout_9)
        self.verticalLayout_10.addLayout(self.verticalLayout_5)
        self.horizontalLayout_13.addWidget(self.groupBox_5)
        self.groupBox_6 = QtWidgets.QGroupBox(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_6.setFont(font)
        self.groupBox_6.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.groupBox_6.setFlat(False)
        self.groupBox_6.setObjectName("groupBox_6")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout(self.groupBox_6)
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout()
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.pushButton_17 = QtWidgets.QPushButton(self.groupBox_6)
        self.pushButton_17.setObjectName("pushButton_17")
        self.verticalLayout_6.addWidget(self.pushButton_17)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_8 = QtWidgets.QLabel(self.groupBox_6)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_10.addWidget(self.label_8)
        self.lineEdit_12 = QtWidgets.QLineEdit(self.groupBox_6)
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.horizontalLayout_10.addWidget(self.lineEdit_12)
        self.verticalLayout_6.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.label_11 = QtWidgets.QLabel(self.groupBox_6)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_11.addWidget(self.label_11)
        self.lineEdit_13 = QtWidgets.QLineEdit(self.groupBox_6)
        self.lineEdit_13.setObjectName("lineEdit_13")
        self.horizontalLayout_11.addWidget(self.lineEdit_13)
        self.verticalLayout_6.addLayout(self.horizontalLayout_11)
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.label_12 = QtWidgets.QLabel(self.groupBox_6)
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_12.addWidget(self.label_12)
        self.lineEdit_14 = QtWidgets.QLineEdit(self.groupBox_6)
        self.lineEdit_14.setObjectName("lineEdit_14")
        self.horizontalLayout_12.addWidget(self.lineEdit_14)
        self.verticalLayout_6.addLayout(self.horizontalLayout_12)
        self.pushButton_19 = QtWidgets.QPushButton(self.groupBox_6)
        self.pushButton_19.setObjectName("pushButton_19")
        self.verticalLayout_6.addWidget(self.pushButton_19)
        self.verticalLayout_11.addLayout(self.verticalLayout_6)
        self.horizontalLayout_13.addWidget(self.groupBox_6)
        self.verticalLayout_12.addLayout(self.horizontalLayout_13)
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.tabWidget.addTab(self.tab_3, "")
        self.verticalLayout_7.addWidget(self.tabWidget)
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName("textBrowser")
        self.verticalLayout_7.addWidget(self.textBrowser)
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        spacerItem = QtWidgets.QSpacerItem(618, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_14.addItem(spacerItem)
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_9.setObjectName("pushButton_9")
        self.horizontalLayout_14.addWidget(self.pushButton_9)
        self.pushButton_10 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_10.setObjectName("pushButton_10")
        self.horizontalLayout_14.addWidget(self.pushButton_10)
        self.verticalLayout_7.addLayout(self.horizontalLayout_14)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.progressBar.sizePolicy().hasHeightForWidth())
        self.progressBar.setSizePolicy(sizePolicy)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_7.addWidget(self.progressBar)
        self.verticalLayout_8.addLayout(self.verticalLayout_7)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 826, 22))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menubar)
        self.menu_2.setObjectName("menu_2")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionopen = QtWidgets.QAction(MainWindow)
        self.actionopen.setObjectName("actionopen")
        self.actionclose = QtWidgets.QAction(MainWindow)
        self.actionclose.setObjectName("actionclose")
        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")
        self.action_3 = QtWidgets.QAction(MainWindow)
        self.action_3.setObjectName("action_3")
        self.actionIEEE754 = QtWidgets.QAction(MainWindow)
        self.actionIEEE754.setObjectName("actionIEEE754")
        self.menu.addAction(self.actionopen)
        self.menu.addSeparator()
        self.menu.addAction(self.action_3)
        self.menu_2.addAction(self.actionAbout)
        self.menu_2.addAction(self.actionIEEE754)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(2)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.groupBox.setTitle(_translate("MainWindow", "大纲转说明-鉴定模板"))
        self.pushButton.setText(_translate("MainWindow", "选择大纲文档"))
        self.label.setText(_translate("MainWindow", "设计人员："))
        self.pushButton_3.setText(_translate("MainWindow", "测试记录转说明"))
        self.pushButton_2.setText(_translate("MainWindow", "开始转换"))
        self.groupBox_2.setTitle(_translate("MainWindow", "追踪关系填写"))
        self.pushButton_4.setText(_translate("MainWindow", "选择文件"))
        self.pushButton_6.setText(_translate("MainWindow", "大纲追踪关系填写（依据大纲文件）"))
        self.pushButton_5.setText(_translate("MainWindow", "测试说明追踪填写（依据说明文件）"))
        self.pushButton_18.setText(_translate("MainWindow", "测试报告追踪填写（依据记录文件）"))
        self.groupBox_4.setTitle(_translate("MainWindow", "根据说明生成测试记录"))
        self.pushButton_11.setText(_translate("MainWindow", "选择文件"))
        self.label_4.setText(_translate("MainWindow", "测试人员："))
        self.label_5.setText(_translate("MainWindow", "监测人员："))
        self.label_6.setText(_translate("MainWindow", "测试时间："))
        self.pushButton_12.setText(_translate("MainWindow", "说明生成记录"))
        self.pushButton_13.setText(_translate("MainWindow", "记录反向生成说明"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "文档生产工具"))
        self.groupBox_3.setTitle(_translate("MainWindow", "UAS单元测试转换"))
        self.label_2.setText(_translate("MainWindow", "被测软件名："))
        self.label_3.setText(_translate("MainWindow", "被测软件标识："))
        self.pushButton_7.setText(_translate("MainWindow", "请选择SAU报告文档"))
        self.pushButton_8.setText(_translate("MainWindow", "开始转换"))
        self.groupBox_5.setTitle(_translate("MainWindow", "自动填充"))
        self.pushButton_16.setText(_translate("MainWindow", "选择文档"))
        self.label_7.setText(_translate("MainWindow", "单元格左侧："))
        self.label_9.setText(_translate("MainWindow", "填充的内容："))
        self.label_10.setText(_translate("MainWindow", "填充的数量："))
        self.pushButton_14.setText(_translate("MainWindow", "清空单元格"))
        self.pushButton_15.setText(_translate("MainWindow", "填充"))
        self.groupBox_6.setTitle(_translate("MainWindow", "提取单元格标题右侧的单元格内容"))
        self.pushButton_17.setText(_translate("MainWindow", "选择文档"))
        self.label_8.setText(_translate("MainWindow", "单元格标题："))
        self.label_11.setText(_translate("MainWindow", "单元格标题："))
        self.label_12.setText(_translate("MainWindow", "单元格标题："))
        self.pushButton_19.setText(_translate("MainWindow", "点击提取"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "文档小工具"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "报告生成工具"))
        self.pushButton_9.setText(_translate("MainWindow", "清空消息"))
        self.pushButton_10.setText(_translate("MainWindow", "显示帮助"))
        self.menu.setTitle(_translate("MainWindow", "文件"))
        self.menu_2.setTitle(_translate("MainWindow", "工具"))
        self.actionopen.setText(_translate("MainWindow", "打开文件"))
        self.actionclose.setText(_translate("MainWindow", "close"))
        self.actionAbout.setText(_translate("MainWindow", "关于工具"))
        self.action_3.setText(_translate("MainWindow", "退出"))
        self.actionIEEE754.setText(_translate("MainWindow", "IEEE转换工具"))
