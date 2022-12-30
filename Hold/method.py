import pandas as pd
from public import *



class Method():

    def fillWithLastDay(dataFrame, start_col=0, start_row=0) -> pd.DataFrame:
        """把所有0的位置按照上一个数补齐"""
        print("Start Position Col: ",start_col, "Row: ",start_row, "DataFrame Shape:", dataFrame.shape)
        for col in range(start_col, dataFrame.shape[1]):
            tem = 0
            for row in range(start_row, dataFrame.shape[0]):
                try:
                    cell = float(dataFrame.iloc[row, col])
                except:
                    cell = dataFrame.iloc[row, col]
                if cell == 0 or cell == 0.0: 
                    dataFrame.iloc[row, col] = tem
                else: 
                    tem = cell
        return dataFrame


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

    def indexCal_1(dataFrame, w) -> pd.DataFrame:

        return