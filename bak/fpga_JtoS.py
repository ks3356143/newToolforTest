from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSignal
from pathlib import *
from PyQt5.QtWidgets import QMessageBox
from docxtpl import DocxTemplate,InlineImage
from docx import Document
import io


class create_FPGA_JtoS(QtCore.QThread):
    sin_out = pyqtSignal(str)

    def __init__(self,parent):
        super().__init__()
        self.parent = parent
        
    def run(self):
        self.sin_out.emit("进入测试记录转说明线程......")
        self.sin_out.emit("开始转换......")
        #如果没有选择文件路径则退出
        if not self.parent.open_file_name:
            self.sin_out.emit('nofile')
            self.parent.tabWidget.setEnabled(True)
            return
        
        #打开模板文件进行渲染，然后就是用docxtpl生成用例
        try:
            tpl_path = Path.cwd() / "need" / "document_templates" / "FPGA记录to说明模板.docx"
            self.sin_out.emit('导入模板文件路径为：' + str(tpl_path))
            tpl = DocxTemplate(tpl_path) #模板导入成功
            
        except:
            QMessageBox.warning(self.parent,"出错了","导入模板出错请检查模板文件是否存在或名字不正确")
            return
        

        try:
            doc = Document(self.parent.open_file_name[0])
            self.sin_out.emit('已识别到FPGA测试记录文件...')
        except:
            self.sin_out.emit('open failed:选择的文档')
            self.parent.tabWidget.setEnabled(True)
            return
        
        self.sin_out.emit('复制测试说明文档模板到本程序所在目录...')
        curpath = Path.cwd() / 'need'
        shuoming_path_tmp = curpath / 'document_templates' / 'FPGA记录to说明模板.docx'
        print(shuoming_path_tmp)
        if shuoming_path_tmp.is_file():
            self.sin_out.emit('已检测到有说明模板文件...')
        else:
            self.sin_out.emit('open failed:选择的文档')
            return
        
        #创建一个字典来储存单个用例
        data_list = []
        #获取表格数量
        tables = doc.tables
        tb_count = len(tables)
        self.sin_out.emit('total:'+ str(tb_count))
        
        for i in range(tb_count):
            if tables[i].cell(0,0).text == '测试用例名称':
                try:
                    data = {'name':'','biaoshi':'','zhuizong':'','zongsu':'','init':'','qianti':'','step':[]}
                    self.sin_out.emit(str(i+1))
                    self.sin_out.emit(f'正在处理第{i+1}个表格')
                    # 1、获取测试用例名称
                    data['name'] = tables[i].cell(0,3).text
                    # 2、获取用例标识
                    data['biaoshi'] = tables[i].cell(0,8).text
                    # 3、获取追踪关系 注意word中换行为\r\x07
                    temp = tables[i].cell(1,3).text
                    data['zhuizong'] = temp.replace("\n", "\r\x07")
                    # 4、获取综述
                    data['zongsu'] = tables[i].cell(2,3).text
                    # 5、初始化
                    data['init'] = tables[i].cell(3,3).text
                    # 6、获取前提与约束
                    data['qianti'] = tables[i].cell(4,3).text
                    # 7、获取步骤信息-总行数减去12为步骤行数 
                    row_count = len(tables[i].rows)
                    step_count = row_count - 12
                    for j in range(step_count):
                        buzhou = {'shuru':'','yuqi':'','num':'','image':'','is_image':'0'}
                        buzhou['num'] = tables[i].rows[7+j].cells[0].text
                        buzhou['shuru'] = tables[i].rows[7+j].cells[2].text
                        cel = tables[i].rows[7+j].cells[2]
                        if len(cel._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'))>0:
                            img_ele = cel._element.xpath('.//pic:pic')[0]
                            embed = img_ele.xpath('.//a:blip/@r:embed')[0]
                            related_part = doc.part.related_parts[embed]
                            image = related_part.image
                            image_bytes = image.blob
                            buzhou['image'] = InlineImage(tpl, io.BytesIO(image_bytes))
                            buzhou['is_image'] = '1'
                        buzhou['yuqi'] = tables[i].rows[7+j].cells[4].text
                        data['step'].append(buzhou)
                    # 8、最后加入data_list
                    data_list.append(data)
                except:
                    self.sin_out.emit(f'第{i}个表格处理错误！')
                    pass
        
        
        #开始渲染模板文件-有2层循环
        try:
            context = {
                "tables":data_list,
                "renyuan":self.parent.lineEdit_17.text(),
            }
            tpl.render(context)
            tpl.save("FPGA反向生成的说明文档.docx")
            self.sin_out.emit('stopsuccess')
        except:
            self.sin_out.emit('stoperror')
            return