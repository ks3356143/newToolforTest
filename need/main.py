# -*- coding: utf-8 -*-
import logging
from unittest import async_case
from warnings import catch_warnings
LOG_FORMAT = "%(asctime)s>%(levelname)s>PID:%(process)d %(thread)d>%(module)s>%(funcName)s>%(lineno)d>%(message)s"
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, )

# 是否打印调试信息标志
debug = True
if debug==True:
    logging.debug("进入主程序，开始导入包...")
    
#导入常规库
import sys,re,string 
from pathlib import *
import shutil
#导入word文档操作库
import win32com
from win32com.client import Dispatch, constants, DispatchEx
import docx
from docxtpl import DocxTemplate
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
        
#关闭线程函数
    def stop_shuoming_thread(self):
        self.create_shuoming_trd.terminate()
        print("停止线程成功！")
    
#~~~~~~~~~~~~~~~~~~~~显示函数~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    def text_display(self, texttmp):
        if texttmp[:10] == 'stopthread':
            self.stop_shuoming_thread()
            return
    
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
#大纲生成测试说明线程
##################################################################################
class create_shuoming(QtCore.QThread):
    sin_out = pyqtSignal(str)

    def __init__(self,parent):
        super().__init__()
        self.parent = parent
        
    def run(self):
        #用来储存测试项DC等转换
        zhuan_dict = {'DC':'文档审查','SU':'功能测试','CR':'代码审查','SA':'静态分析','AC':'静态测试',\
            'IO':'接口测试','SE':'安全性测试','BT':'边界测试','RE':'恢复性测试','ST':'强度测试',\
                'AT':'余量测试','GUI':'人机交互界面测试','DP':'数据处理测试','JR':'兼容性测试',\
                    'LG':'逻辑测试'}
        
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
        #self.w.visible=True
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
        
        #创建一个字典来储存单个用例
        
        data_list = []
        
        #获取表格数量
        try:
            csx_tb_count = dagangfile.Tables.Count
            self.sin_out.emit('total:'+ str(csx_tb_count))
        except:
            self.sin_out.emit('不存在表格！')
            QMessageBox.warning(self.parent,'出错了','测试说明文档格式错误或者没有正确表格')
            dagangfile.Close()
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
            return
        
        #循环表格
        yongli_count = 0
        #用来储存章节号中的DC、SU等标识，用于章节号判断
        is_fire_su = ""
        #用来储存基本_分割后个数
        num_fenge = 3
        for i in range(csx_tb_count):
            self.sin_out.emit(str(i))
            if dagangfile.Tables[i].Rows.Count > 2:
                #注意win32com的Cell从1开始不是从0开始
                if dagangfile.Tables[i].Cell(1, 1).Range.Text.find('测试项名称') != -1:               
                    #一个用例不变内容获取
                    dagangfile.Tables[i].Rows.First.Select()
                    zhangjiehao = self.w.Selection.Bookmarks("\headinglevel").\
                        Range.Paragraphs(1).Range.ListFormat.ListString
                    zhangjieming = self.w.Selection.Bookmarks("\headinglevel").\
                        Range.Paragraphs(1).Range.Text.rstrip('\r')
                    print("测试项所在章节号：",zhangjiehao)
                    #获取用例标识不加上序号_1
                    basic_biaoshi = dagangfile.Tables[i].Cell(1,4).Range.Text.rstrip()[:-2]
                    print("测试项标识为：",basic_biaoshi)
                    
                    #储存num_fenge的数值，初始化
                    if yongli_count == 0:
                        num_fenge = len(basic_biaoshi.split("_"))
                    #获取测试用例名称Cell(4,2)整行
                    info_ceshi_buzhou = dagangfile.Tables[i].Cell(4,2).Range.Text
                    info_ceshi_yuqi = dagangfile.Tables[i].Cell(9,2).Range.Text
                    #判断是否只有一行，如果只有一行处理表格
                    if dagangfile.Tables[i].Cell(4, 2).Range.Paragraphs.Count <= 1:
                        
                        #缓存一个用例的data填入数据
                        data = {'zhangjie':'','mingcheng':'','biaoshi':'','is_first':'1', \
                        'yueshu':'软件正常工作，环境连接正常', 'yongli_biaoshi':'','renyuan':'陈俊亦',\
                        'chushi':'外接设备或软件运行正常','csx_mingcheng':'','is_begin':'0',\
                            'zongsu':'',"zuhe":[],'csxbs':""}
                        zuhe_dict = {"buzhou":"","yuqi":"","xuhao":"1"}
                        try:
                            #填写一行情况下表格
                            data['mingcheng'] = dagangfile.Tables[i].Cell(1,2).Range.Text.rstrip('\r\x07')
                            #注意word中后面都有2个字符
                            data['yongli_biaoshi'] = basic_biaoshi + "_1"
                            data['zhangjie'] = zhangjiehao
                            data['csx_mingcheng'] = zhangjieming
                            data['biaoshi'] = basic_biaoshi
                            data['zongsu'] = dagangfile.Tables[i].Cell(3,2).Range.Text[:-2]
                            
                            zuhe_dict["buzhou"] = dagangfile.Tables[i].Cell(4,2).Range.Text.rstrip('\r\x07')
                            zuhe_dict["yuqi"] = dagangfile.Tables[i].Cell(9,2).Range.Text.rstrip('\r\x07')
                            zuhe_dict["xuhao"] = '1'
                            data["zuhe"].append(zuhe_dict)
                            #判断是否为第一个测试类型，如果是修改章节号标识,则将章节号展示，如果和储存相同
                            #则章节号不展示
                            #首先获取测试项标识，分割成列表
                            fenge = data['biaoshi'].split("_")
                            #获取当前测试项分割后的个数
                            if len(fenge) < num_fenge:
                                if fenge[-1]!= is_fire_su:
                                    is_fire_su = fenge[-1]
                                    data['is_begin'] = "1"
                                    data['csxbs'] = zhuan_dict[fenge[-1]]
                            else:
                                if fenge[-2] != is_fire_su:
                                    is_fire_su = fenge[-2]
                                    data['is_begin'] = "1"
                                    data['csxbs'] = zhuan_dict[fenge[-2]]
                            
                            #将data加入data_list
                            data['is_first'] = "1"
                            data_list.append(data)
                            yongli_count += 1
                            
                            
                            self.sin_out.emit('###获取用例序号：{}'.format(yongli_count))
                            self.sin_out.emit('###该用例标识为：{}'.format(data['yongli_biaoshi']))
                        except:
                            QMessageBox.warning(self.parent,"出错了","获取文档信息出错检查文档格式")
                        
                    elif dagangfile.Tables[i].Cell(4, 2).Range.Paragraphs.Count > 1:
                        
                        try:
                            #下面拆分每行，使用\r（回车）分割
                            info_buzhou_list = list(filter(lambda x:x!="\x07" and x!="",info_ceshi_buzhou.split('\r')))
                            info_yuqi_list = list(filter(lambda x:x!="\x07" and x!="",info_ceshi_yuqi.split('\r')))
                            
                            #去掉括号和以下字符
                            rule = "[()（）;；。]" #rule为去掉的符号（这个可以改TODO）

                            #初始化去掉rule的列表
                            buzhou_list = []
                            yuqi_list = []
                            
                            for item in info_buzhou_list:
                                buzhou_list.append(re.sub(rule,"",item).strip(string.digits))
                            for item in info_yuqi_list:
                                yuqi_list.append(re.sub(rule,"",item).strip(string.digits))
                            
                            #获取测试项综述-为该循环前不变内容
                            basic_zongshu = buzhou_list.pop(0).strip()
                            print('获取的测试用例综述是：',data["zongsu"])
                            if len(buzhou_list) == len(yuqi_list):
                                self.sin_out.emit('检测到格式预期和步骤的行数正确...')
                            
                            #获取字典中的buzhou和yuqi,找冒号
                            j = -1 #自制列表索引
                            substrict_list = [] #差值列表
                            for item in buzhou_list:
                                #先找到冒号所在索引
                                j = j + 1
                                if item.find(":") != -1 or item.find("：") != -1:
                                    #现在知道冒号所在行号，现在要确定每个用例几行
                                    substrict_list.append(j)
                                    
                            #！！！注意差值计算步骤需要-1才是正确的步骤数量
                            self.sin_out.emit("解析测试项序号"+ str(i) + "|检测到冒号所在行号为：" \
                                + str(substrict_list))
                            self.sin_out.emit("|检测到步骤总行数(序号)：" \
                                + str(len(buzhou_list)))
                            
                            #循环用例个数
                            count_test = len(substrict_list)
                            temp_list = substrict_list
                            temp_list.append(len(buzhou_list))
                            
                            for item in range(count_test):
                                #初始化data数据,包括步骤和预期、序号dict
                                data = {'zhangjie':'','mingcheng':'','biaoshi':'','is_first':'0', \
                                    'yueshu':'软件正常工作，环境连接正常', 'yongli_biaoshi':'','renyuan':'陈俊亦',\
                                    'chushi':'外接设备或软件运行正常','csx_mingcheng':'','is_begin':'0',\
                                        'zongsu':'',"zuhe":[],'csxbs':""}
                                
                                #这里要求冒号最后一个
                                data['mingcheng'] = buzhou_list[substrict_list[item]][:-1]
                                data['yongli_biaoshi'] = basic_biaoshi + f'_{item+1}'
                                #常规填入
                                data['zhangjie'] = zhangjiehao
                                data['csx_mingcheng'] = zhangjieming
                                data['biaoshi'] = basic_biaoshi
                                data['zongsu'] = basic_zongshu
                                #步骤填入,首先根据substrict_list获取有几个步骤
                                
                                #循环行数
                                for x in range(temp_list[item+1]-(temp_list[item]+1)): #循环一个用例步骤预期数
                                    zuhe_dict = {"buzhou":"","yuqi":"","xuhao":""}
                                    #把每个步骤和预期都放进zuhe_dict中
                                    zuhe_dict["buzhou"] = buzhou_list[temp_list[item]+x+1]
                                    zuhe_dict["yuqi"] = yuqi_list[temp_list[item]+x+1]
                                    zuhe_dict["xuhao"] = str(x+1)
                                    data["zuhe"].append(zuhe_dict)
                                    
                                if item == 0:
                                    data['is_first'] = '1'
                                #判断是否为SU标题
                                fenge = data['biaoshi'].split("_")
                                #获取当前测试项分割后的个数
                                if fenge[-2] != is_fire_su:
                                    is_fire_su = fenge[-2]
                                    data['is_begin'] = "1"
                                    data['csxbs'] = zhuan_dict[fenge[-2]]
                                    
                                #加入data_list
                                data_list.append(data)
                                yongli_count += 1 #用例计数加一
                                
                                self.sin_out.emit('###获取用例序号：{}'.format(yongli_count))
                                self.sin_out.emit('###该用例标识为：{}'.format(data['yongli_biaoshi']))
                        except:
                            QMessageBox.warning(self.parent,'出错了','请检查大纲格式如（冒号）（行数）！，\
                                这会导致一个测试项生成失败，但不影响其他测试项有运行！')

                    else:
                        QMessageBox.warning(self.parent,'出错了','请检查大纲行数是否按要求')
                        continue
                    
            else:
                QMessageBox.warning(self.parent,'出错了','请检查表格格式')
            
            #调试查看整体data_list
            #self.sin_out.emit(str(data_list)) 
        #关闭大纲文档（因为以及提取完毕）
        try:
            dagangfile.Close()
            self.w.Quit()
            pythoncom.CoUninitialize()
            #处理完毕结束语
            #self.sin_out.emit('function success')
            #self.sin_out.emit('请参考当前文件夹下的《测试说明模板.docx》，记得及时更改文件名或者复制备份，防止被覆盖。')
        except:
            self.sin_out.emit('function fail')
            self.w.Quit()
            pythoncom.CoUninitialize()
            return
    
        #打开模板文件进行渲染，然后就是用docxtpl生成用例
        try:
            tpl_path = Path.cwd() / "need" / "document_templates" / "测试说明模板.docx"
            self.sin_out.emit('导入模板文件路径为：' + str(tpl_path))
            tpl = DocxTemplate(tpl_path) #模板导入成功
            
        except:
            QMessageBox.warning(self.parent,"出错了","导入模板出错请检查模板文件是否存在或名字不正确")
            return
        
        #开始渲染模板文件-有2层循环
        try:
            context = {
                "tables":data_list,
            }
            tpl.render(context)
            tpl.save("生成的说明文档.docx")
            QMessageBox.warning(self.parent,"生成文档成功","请查看当前工具根目录（生成的说明文档.docx）")
            self.sin_out.emit('stopthread')
        except:
            QMessageBox.warning(self.parent,"生成文档出错","生成文档错误，请确认模板文档是否已打开或格式错误")
            self.sin_out.emit('stopthread')
            return
        