import os
import re
import sys
current_dir = os.path.abspath(os.path.dirname(__file__)+'/'+'..')
sys.path.append(current_dir)
# import sys
# current_dir = os.path.abspath(os.path.dirname(__file__))
# sys.path.append(current_dir)
# sys.path.append("..")
from usermodules.MyException import MyException
def get_files(floder_path):
    """
    获取路径下所有文件
    :param source_path: 文件夹路径
    :return: 获取文件名字,如果没有路径则报错
    """
    if not exist_file(floder_path):
        raise MyException.FileErrorException('no file found,please check your floder_path')
    return [os.path.join(floder_path,file)for file in next(os.walk(floder_path))[2]]


def get_folder_names(floder_path):
    """
    获取路径下所有文件夹
    :param floder_path: 文件夹路径
    :return: 获取文件上级目录,如果没有路径则报错
    """
    if not exist_file(floder_path):
        raise MyException.FileErrorException('no file found,please check your floder_path')
    return [os.path.join(floder_path,file)for file in next(os.walk(floder_path))[1]]


def get_files_by_extension(floder_path='.', extension_name='pdf'):
    """
    寻找目录下指定格式的文件
    :floder_path 需要搜索的目录:
    :extension_name 文件扩展名字符串:
    :return: 文件列表
    """
    if not exist_file(floder_path):
        raise MyException.FileErrorException('no file found,please check your floder_path')
    pattern=r'.+?\.'+extension_name
    filename=get_files(floder_path)
    return [os.path.join(floder_path, file) for file in filename if re.match(pattern, file)]


def split_path(file_path):
    """
    分割文件路径
    :param file_path: 文件绝对路径
    :return: 父文件夹路径，文件名，扩展名
    示范：
    file_path="D:/test/test.py"
    father_path为文件的目录,即D:/test
    file_name为文件的名字,即test
    file_extension为文件的扩展名,即.py
    """
    father_path,file_name_by_extension=os.path.split(file_path)
    file_name,file_extension=os.path.splitext(file_name_by_extension)
    return father_path,file_name,file_extension


def get_curdir():
    """
    :return: 获取当前目录
    """
    return os.getcwd()
def check_extension(file_path,extension_name='pdf'):
    """

    :param file_path: 文件路径
    :param extension_name: 扩展名
    :return: 如果符合扩展名则返回True
    """
    pattern = r'.+?\.' + extension_name
    if re.match(pattern,file_path):
        return True
    return False
def exist_file(file_path):
    return os.path.exists(file_path)
