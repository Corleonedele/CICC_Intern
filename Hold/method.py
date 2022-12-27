import pandas as pd
from public import *



class Method():

    def codeTransfer():
        return

    def mergeSameSheet(file_name, sheet_name, path, **kwargs) -> pd.DataFrame:
        """合并同一个文件夹下多个Excel中的指定sheet"""
        file_list = Public.getFileList(filter_arg=file_name, path=path)
        print("Current Files:", file_list)
        if len(file_list)==0:
            return ValueError
        else:
            return_sheet=pd.DataFrame()

        for file in file_list:
            file_path = path + "/" + file
            read_sheet =  pd.read_excel(file_path, header=kwargs["header"], index_col=kwargs["index_col"], sheet_name=sheet_name)
            return_sheet=pd.concat([return_sheet, read_sheet], ignore_index = False)
        return return_sheet.groupby(return_sheet.index).first() 

    def fillEmptyCell(fill_value=0, fill_position="MAX") -> pd.DataFrame():
        """改函数用于填补所有Excel中的空缺, fill_value代表填充的值, fill_position代表填充到哪个位置"""
        pass
