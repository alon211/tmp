import sys
import os
current_dir = os.path.abspath(os.path.dirname(__file__)+'/'+'..')
sys.path.append(current_dir)
FILE_ERROR_INFOR={
    'FILE_NOT_FOUNT':"file can not founded,please check your path",
    'FILE_EXTESION_ERROR':'文件扩展名不匹配'
}
class MyException(Exception):
    def __int__(self,*args):
        self.args=args
class FileErrorException(MyException):
    def __int__(self,message):
        self.message=message