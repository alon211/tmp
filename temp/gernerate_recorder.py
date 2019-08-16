import pandas as pd
import win32com.client as win32
import sys
from datetime import datetime
import os
import shutil
import pysnooper
import traceback
import copy


class recoder_process:
    def __init__(self,t_path=None,export_path=None):
        if export_path is not None:
            self._ex_path=export_path
        else:
            self._ex_path=None
        if t_path is not None:
            self._template_path=t_path
        else:
            self._template_path=None
        
        self._exception=False,''
        try:
            self._excel=win32.gencache.EnsureDispatch('Excel.Application')
        except:
                self._exception = True, '无法获取excel应用程序，请确认已经安装了office excel'
        self._wb=None
        self._ws=None

        self._startrow=2
        self._cur_startrow=self._startrow
        self._startcolumn='A'
        self._cur_startcolumn = self._startcolumn
        self._startcolumnMax='O'

        self._endrow = 9
        self._cur_endrow = self._endrow
        self._endcolumn = 'G'
        self._cur_endcolumn = self._endcolumn
        self._endcolumnMax='U'
        self._init_data=[
                            ['产品名称:1型监控屏', None, None, '序列号:      ', None, '版本:', None],
                            ['部件名称', '供货商', '识别号', '部件版 本', 'PCB版本', '序列号', '备注'],
                            ['电源板', 'SRI', '701050000477100', None, None, None, None],
                            ['CPU板', '研华', '500001852710002', None, None, None, None],
                            ['液晶屏', '京东方', '500000315440005', '/', None, None, None],
                            ['AD板', '京东方', '500000116710000', '/', None, None, None],
                            ['驱动板', '京东方', '500000103620003', '/', None, None, None],
                            ['触摸屏', '京东方', '500000310410002', '/', None, None, None]]
        self._template_data=None  #模板一页的空数据只保留格式
        self._current_data=copy.deepcopy(self._init_data)
        self._all_data=[]
    def init_data(self):
        self._cur_endrow = self._endrow
        self._cur_startrow = self._startrow
        self._cur_startcolumn = self._startcolumn
        self._cur_endcolumn = self._endcolumn
        self._current_data = copy.deepcopy(self._init_data)
    def add_current_data(self):
        self._all_data.append(self._current_data)
    @property
    def all_data(self):
        return self._all_data
    @all_data.setter
    def all_data(self,value):
        self._all_data=value
    @property
    def template_path(self):
        return self._template_path
    @template_path.setter
    def template_path(self,value):
        self._template_path=value
    @property
    def ex_path(self):
        return self._ex_path
    @ex_path.setter
    def ex_path(self,value):
        self._ex_path=value
    
    @property
    def current_data(self):
        return self._current_data
    @current_data.setter    
    def current_data(self,value):
        self._current_data=value
        
    @property
    def exception(self):
        return self._exception

    @pysnooper.snoop('debug.log', depth=2)
    def get_excel_obj(self,path):
        """

        :param path: 模板表格路径
        :return: 返回模板表格对象
        """

        try:
            if self._wb is not None:
                self._wb.Close()
            self._wb = self._excel.Workbooks.Open(path)
            self._excel.Visible = False
            self._ws=self._wb.Worksheets('记录表')
            self._template_data=self._ws.Range('A1:G44').Value
            self._excel.CutCopyMode = False
        except Exception as e:
            print(e)
            self._exception = True, '模板文件对象获取失败,请确认模板文件是否存在！'

    @pysnooper.snoop('debug.log', depth=2)
    def create_new_file(self):

        file = f'{datetime.now().strftime("%Y-%m-%d %H %M %S")}.xls'
        self.init_data()
        if self._ex_path is not None:
            shutil.copyfile(self._template_path,os.path.join(self._ex_path,file))
        else:
            self._exception = True, '输出文件路径不存在'
        self.get_excel_obj(os.path.join(self._ex_path,file))
    def read_recordersheet(self):
        return self._ws.Range("A1:G9").Value

    @pysnooper.snoop('debug.log', depth=2)
    def copy_recoderformate(self,startcolumn:str,endcolumn:str):
        self._ws.Range(f"A1:G44").Copy()
        target=self._ws.Range(f"{startcolumn}1:{endcolumn}44")
        target.PasteSpecial(8,-4142)
        target.PasteSpecial(13,-4142)
        # 黏贴空数据
        target.Value=self._template_data
    # def init_currentData(self):
    #     self._current_data=self._init_data
    @pysnooper.snoop('debug.log', depth=2)
    def write_recordersheet(self,data):
        '''

        :param data: 修改的数据
        :return:
        '''
        try:
            
            r=f'{self._cur_startcolumn}{self._cur_startrow}:{self._cur_endcolumn}{self._cur_endrow}'
            self._ws.Range(r).Value=data
            if self._cur_startrow>=29:
                # 如果第一页填满了换页
                self._cur_startrow=self._startrow
                self._cur_endrow=self._endrow
                if self._cur_startcolumn>='O':
                    #每个文件最多三页,三页后新建一个文件
                    self.savefile()
                    self.create_new_file()
                    return
                self._cur_startcolumn=chr(ord(self._cur_startcolumn)+7)
                self._cur_endcolumn=chr(ord(self._cur_endcolumn)+7)
                # 新建一页空数据
                self.copy_recoderformate(self._cur_startcolumn,self._cur_endcolumn)

            else:
                self._cur_startrow+=9
                self._cur_endrow +=9
            print(f'{self._cur_startcolumn}{self._cur_startrow}:{self._cur_endcolumn}{self._cur_endrow}')
        except:
            a = traceback.format_exc(limit=1)
    def savefile(self):
        self._wb.Save()
    def quit(self):
        self.savefile()
        self._excel.Application.Quit()

# a=recoder_process('E:\myproject\智能工作台',r'E:\myproject\智能工作台\1.xls')
# if a.exception[0]:
#     print(a.exception[1])
#     sys.exit(0)
# a.create_new_file()
# if a.exception[0]:
#     print(a.exception[1])
#     a.quit()
#     sys.exit(0)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)
# a.write_recordersheet(None)

# a.quit()

