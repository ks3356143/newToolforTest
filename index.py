# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtCore,QtGui
from need.main import userMain
import qtmodern.styles
import qtmodern.windows
# from qt_material import apply_stylesheet
#以下导入为打包导入所需-使用软件
import json
import docx
import docxtpl
import six
import docxcompose
import lxml
import markupsafe
import win32api
import win32com

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    # 设置任务栏软件图标
    app.setWindowIcon(QtGui.QIcon('Icon.png'))
    win = userMain()
    ##以下是qt_material样式加载
    # apply_stylesheet(app,theme = 'dark_teal.xml')
    # win.show()
    
    qtmodern.styles.light(app) #还有dark可以选择
    mw = qtmodern.windows.ModernWindow(win)
    mw.show()
    '''
    #设置窗口有边框可拖动，但删除标题栏
    self.setWindowFlags(
    Qt.Window | Qt.CustomizeWindowHint | Qt.WindowSystemMenuHint)
    # win.show()
    '''
    sys.exit(app.exec_())