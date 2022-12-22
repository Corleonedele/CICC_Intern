import pandas as pd
import numpy as np
import os


"""全局参数"""

class Public():

    def initDirs():
        """初始化文件系统
        Fixed_Data 用于存储每日需要用到的固定文件
        Input_Data 用于添加每日需要用到的新增文件
        Output_Data 用于储存程序运行完的文件
        Stored_Data 用于备份各个文件
        """
        if not os.path.exists("./Fixed_Data"):
            os.makedirs("./Fixed_Data") 
        if not os.path.exists("./Input_Data"):
            os.makedirs("./Input_Data") 
        if not os.path.exists("./Output_Data"):
            os.makedirs("./Output_Data") 
        if not os.path.exists("./Stored_Data"):
            os.makedirs("./Stored_Data") 

    def getFileList(filter_arg=False, path=".", dir_only=False):
        """默认获取当前文件夹下文件, 可设置filter_arg来查找特定文件(eg filter_arg=".py" 会查找python文件, 设置path去查找特定路径下查找文件, 返回该文件夹下所有文件"""
        if filter_arg:
            return list(filter(lambda x: filter_arg in x, os.listdir(path)))
        else:
            if dir_only:
                return list(filter(lambda x: "." not in x, os.listdir(path)))
            return os.listdir(path)
            
    def readFile(file_name):
        """按文件名查找文件, 如果当前目录没找到, 找且仅找下一级目录, 若仍然没找到则返回None"""
        current_file_list = Public.getFileList()
        if file_name in current_file_list: 
            return os.path.join("./", file_name)
        else:
            for dir in Public.getFileList(dir_only=True):
                if file_name in Public.getFileList(path=dir):
                    return os.path.join("./"+dir, file_name)
        return None


class Compute():
    def maxDrawdown(input_list) -> list:
        """计算最大回撤 返回最大回撤率与最大回撤区间(左, 右)的index"""
        i = np.argmax((np.maximum.accumulate(input_list) - input_list) / np.maximum.accumulate(input_list))
        if i == 0: 
            return 0
        j = np.argmax(input_list[:i])
        return (input_list[j] - input_list[i]) / input_list[j], j, i

