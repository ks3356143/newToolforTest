import pythoncom
from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSignal
from win32com.client import DispatchEx
from pathlib import *
from PyQt5.QtWidgets import QMessageBox
from docxtpl import DocxTemplate

class create_bujian(QtCore.QThread):
    sin_out = pyqtSignal(str)

    def __init__(self,parent):
        super().__init__()
        self.parent = parent
        
    def run(self):
        self.sin_out.emit("进入部件测试获取调用函数线程......")
        self.sin_out.emit("开始填写文档......")
         
        #如果没有选择路径则退出
        if not self.parent.open_file_name:
            self.sin_out.emit('nofile')
            self.parent.tabWidget.setEnabled(True)
            return
        
        #告诉windows单线程
        pythoncom.CoInitialize()
        #在用户选择的目录中查找UAS单位测试报告文档
        self.sin_out.emit('打开单元测试原文件...')
        
        #使用win32com打开-记得关闭
        #打开word应用
        self.w = DispatchEx('Word.Application')
        #self.w.visible=True
        self.w.DisplayAlerts = 0
        try:
            bujianfile = self.w.Documents.Open(self.parent.open_file_name[0])
        except:
            self.sin_out.emit('open failed:选择的文档')
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
            return
        
        curpath = Path.cwd() / 'need'
        danyuan_file_tmp = curpath / 'document_templates' / '部件桩函数工具1.docx'
        print(danyuan_file_tmp)
        
        if danyuan_file_tmp.is_file():
            self.sin_out.emit('已检测到有部件模板文件...')
        else:
            self.sin_out.emit('open failed:选择的文档')
            return
        
        #创建个列表放数据-important
        data_list = []
        
        #try统计表格数量
        try:
            csx_tb_count = bujianfile.Tables.Count
            self.sin_out.emit('total:'+ str(csx_tb_count))
            self.sin_out.emit("正在调用word文档操作接口,可能会有点慢...")
        except:
            self.sin_out.emit('不存在表格！')
            QMessageBox.warning(self.parent,'出错了','测试说明文档格式错误或者没有正确表格')
            try:
                bujianfile.Close()
            except:
                pass
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
            return
        
        #开始处理表格-important
        #我先统计有多少个生成的表格-即用例有多少个呗
        yongli_count = 0
        for i in range(csx_tb_count):
            if bujianfile.Tables[i].Rows.Count > 2:
                #注意win32com的Cell从1开始不是从0开始
                if bujianfile.Tables[i].Cell(1, 1).Range.Text.find('用例名称') != -1: 
                    yongli_count += 1
        
        
        yongli_num = 0
        hanshuming_duibi = ''
        alowFunctionInject = True
        for i in range(csx_tb_count):
            self.sin_out.emit(str(i))
            #准备填入的data
            data = {'functionName':'','subitem':[]}
            
            #找到函数名,这里容易出问题~~~~~~~~~~~~~~~~
            if bujianfile.Tables[i].Rows.Count > 2:
                if bujianfile.Tables[i].Cell(1, 1).Range.Text.find('功能描述') != -1: 
                    bujianfile.Tables[i].Cell(1, 1).Range.Select()
                    self.w.Selection.MoveUp()
                    self.w.Selection.MoveUp()
                    self.w.Selection.MoveUp()
                    s = self.w.Selection.Paragraphs(1).Range.Text[:-1]
                    s1 = s.split(". ")[-1]
                    #放入函数名比对
                    data['functionName'] = s1
                    data_list.append(data)
                    yongli_num += 1 #用例创建加一

            #找章节号~~~~~~~~~~~~~~~~~~~~~~~~
            if bujianfile.Tables[i].Rows.Count > 2:
                if bujianfile.Tables[i].Cell(1, 1).Range.Text.find('用例名称') != -1:  
                    #函数名获取
                    if s1 != hanshuming_duibi:
                        hanshuming_duibi = s1
                        alowFunctionInject = True
                    else:
                        alowFunctionInject = False
                            
                elif bujianfile.Tables[i].Cell(1, 2).Range.Text.find('定义') != -1: 
                    #定义个桩函数dict
                    if alowFunctionInject == True:
                        temp = bujianfile.Tables[i].Cell(1, 3).Range.Text[:-2]
                        temp1 = temp.split("(")[0]
                        temp2 = temp1.split(" ")[-1]
                        data_list[yongli_num - 1]['subitem'].append(temp2)
        
        
        print('最后data_list',data_list)    
        #最后关闭文档
        try:
            self.w.Quit()
            pythoncom.CoUninitialize()
            self.parent.tabWidget.setEnabled(True)
        except:
            QMessageBox.warning(self.parent,"关闭文档失败","关闭文档失败！")
            return
        
        try:
            tpl_path = Path.cwd() / "need" / "document_templates" / "部件桩函数工具1.docx"
            self.sin_out.emit('导入模板文件路径为：' + str(tpl_path))
            tpl = DocxTemplate(tpl_path) #模板导入成功
        except:
            QMessageBox.warning(self.parent,"出错了","导入模板出错请检查模板文件是否存在或名字不正确")
            return
        
        #开始渲染模板文件
        try:
            context = {
                "tables":data_list,
            }
            tpl.render(context)
            tpl.save("部件提起调用函数表格.docx")
            self.sin_out.emit('stopsuccess')
        except:
            self.sin_out.emit('stoperror')
            return