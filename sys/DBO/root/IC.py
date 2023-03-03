
import openpyxl
from public import Public


class IC():

    path = "./DBO/DB/IC/IC数据测试.xlsx"

    def 比例计算():
        wb = openpyxl.open(IC.path)
        比例数据st = wb["IC比例数据"]
        交易记录st = wb["交易记录"]
        product_code = []
        
        for row in range(1, 交易记录st.max_row):
            product_code.append(交易记录st.cell(row, 3).value)
            
        pass

IC.比例计算()