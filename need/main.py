# -*- coding: utf-8 -*-
import logging
LOG_FORMAT = "%(asctime)s>%(levelname)s>PID:%(process)d %(thread)d>%(module)s>%(funcName)s>%(lineno)d>%(message)s"
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, )

# 是否打印调试信息标志
debug = True
if debug==True:
    logging.debug("进入主程序，开始导入包...")
    
#导入常规库
import sys
import os
from pathlib import *
import shutil
#导入word文档操作库
import win32com
from win32com.client import Dispatch, constants, DispatchEx
import docx
import docxtpl
import pythoncom
#导入QT组件
from PyQt5 import QtCore,QtGui
from PyQt5.QtGui import QIntValidator
from PyQt5.QtCore import QTranslator
from PyQt5.QtWidgets import *
from PyQt5.QtCore import pyqtSignal
# #导入数据处理库
# import pandas as pd

#导入UI转换PY文件
from need.Ui_GUI import Ui_MainWindow
#导入工具包文件-时间转换
from need.utils import get_current_time,get_current_name,get_current_date,get_current_hour

class userMain(QMainWindow,Ui_MainWindow):
    #自定义信号和槽
    
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        if debug == True:
            logging.debug("初始化主程序:")
        
        
        # 实例化翻译家
        self.trans = QTranslator()
        self.setWindowTitle('测试个人工具')
        
        #存放文件夹路径变量
        self.open_dirs_name = ''
        #存放文件名称路径变量
        self.open_file_name = ''
        
        #读取配置文件
        
        #连接线程函数
        ##连接大纲生成说明函数
        self.create_shuoming_trd = create_shuoming(self) #生成说明的线程
        self.create_shuoming_trd.sin_out.connect(self.text_display) #信号绑定输出的区域
        self.pushButton_2.clicked.connect(self.create_shuoming_btn) #点击按钮执行线程
        
        #自定义信号连接
        
        # 获取状态栏对象
        self.user_statusbar = self.statusBar()
        # 右下角窗口尺寸调整符号
        self.user_statusbar.setSizeGripEnabled(False)
        self.user_statusbar.setStyleSheet("QStatusBar.item{border:10px}")
        
        #~~~~~~~~~~~~~·按钮连接函数~~~~~~~~~~~~~~~~
        ##选择文件按钮连接
        self.pushButton.clicked.connect(self.choose_docx_func)
        
#~~~~~~~~~~~~~~~~~~~~初始化直接运行的函数（也就是起始运行一次）~~~~~~~~~~~~~~~~~~~~~~~~~~


#~~~~~~~~~~~~~~~~~~~~间接按钮函数，用户点击后操作~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



#~~~~~~~~~~~~~~~~~~~~线程区域函数，用于启动线程~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#生成测试说明启动函数
    def create_shuoming_btn(self):
        self.create_shuoming_trd.start()
        return
    
#选择文档函数
    def choose_docx_func(self):
        self.open_file_name = QFileDialog.getOpenFileName(
            self, '选择文件', '.', "Word files(*.docx)")
        self.textBrowser.append('已选择文件路径：' + self.open_file_name[0])
    
#~~~~~~~~~~~~~~~~~~~~显示函数~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    def text_display(self, texttmp):
            if texttmp[:6] == 'total:':
                cnt = int(texttmp[6:])
                self.progressBar.setRange(0, cnt - 1)

            if texttmp == 'function success':
                QMessageBox.information(self, '操作成功', '请查看本程序当前文件夹下的相关文档！')
                self.textBrowser.append('完成！！！')
                self.tabWidget.setEnabled(True)
                return
            if texttmp == 'function fail':
                QMessageBox.warning(self, '出错了', '保存文件失败！')
                self.tabWidget.setEnabled(True)
                return
            if texttmp == 'no folder':
                QMessageBox.information(self, '没有选择文件夹', '还没有选择文件夹，点击"文件"菜单进行选择！')
                return
            if texttmp.find('warning:') != -1:
                QMessageBox.information(self, 'WARNING', texttmp[8:])
                return

            if texttmp.find('open failed:') != -1:
                QMessageBox.warning(
                    self, '打开文件失败',
                    '打开' + texttmp[12:] + '失败' + '请确认文档是否打开或者模板文件存在且后缀名为docx！')
                return
            if texttmp == 'copy failed':
                QMessageBox.warning(self, '复制文件失败', '复制文件失败了，注意原文件不要放在本程序根目录下！')
                return
            if texttmp == 'nofile':
                QMessageBox.information(self, '错误',
                                        '还没有选择文件（夹），点击"文件"菜单或者工具栏进行选择！')
                return
            if texttmp.isdigit() == True:
                self.progressBar.setValue(int(texttmp))
            else:
                self.textBrowser.append(texttmp)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '提示',
                    "是否要关闭所有窗口?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
            sys.exit(0)   # 退出程序
        else:
            event.ignore()
    
##################################################################################
#
#大纲生成测试说明线程
#
##################################################################################
class create_shuoming(QtCore.QThread):
    sin_out = pyqtSignal(str)

    def __init__(self,parent):
        super().__init__()
        self.parent = parent
        
    def run(self):
        self.sin_out.emit("进入军品大纲转说明......")
        self.sin_out.emit("开始转换......")
        #如果没有选择路径则退出
        if not self.parent.open_file_name:
            self.sin_out.emit('nofile')
            self.parent.tabWidget.setEnabled(True)
            return
        #告诉windows单线程
        pythoncom.CoInitialize()
        #在用户选择的目录中查找大纲文档
        self.sin_out.emit('打开测评大纲文档...')
        
        #使用win32com打开-记得关闭
        #打开word应用
        self.w = DispatchEx('Word.Application')
        #w.visible=0
        self.w.DisplayAlerts = 0
        try:
            dagangfile = self.w.Documents.Open(self.parent.open_file_name[0])
        except:
            self.sin_out.emit('open failed:选择的文档')
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
            return
        
        self.sin_out.emit('复制测试说明文档模板到本程序所在目录...')
        curpath = Path.cwd() / 'need'
        shuoming_path_tmp = curpath / 'document_templates' / '测试说明模板.docx'
        print(shuoming_path_tmp)
        if shuoming_path_tmp.is_file():
            self.sin_out.emit('已检测到有说明模板文件...')
        else:
            self.sin_out.emit('open failed:选择的文档')
            return
        
        #复制模板文件到根目录
        shutil.copy(shuoming_path_tmp, curpath.parent)
        #创建一个字典来储存单个用例
        data = {'zhangjie':'','mingcheng':'','biaoshi':'','is_first':'', \
            'yueshu':'', \
            'chushi':'外接设备或软件运行正常','csx_mingcheng':'','zongsu':'',"buzhou":[],\
                'yuqi':[]}
        data_list = []
        
        #获取表格数量
        try:
            csx_tb_count = dagangfile.Tables.Count
        except:
            self.sin_out.emit('no table')
            #QMessageBox.warning(self,'出错了','测试说明文档格式错误或者没有正确表格')
            dagangfile.Close()
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
            return
        
        #循环表格
        for i in range(csx_tb_count):
            if dagangfile.Tables[i].Rows.Count > 2:
                #注意win32com的Cell从1开始不是从0开始
                if dagangfile.Tables[i].Cell(1, 1).Range.Text.find('测试项名称') != -1:
                    #获取章节号
                    
            else:
                continue
                
                
        
        
        
        return