# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtCore
from need.main import userMain
import qtmodern.styles
import qtmodern.windows
#以下导入为打包导入所需
import json
import docx
import docxtpl
import six
import docxcompose
import lxml
import markupsafe

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    win = userMain()

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