from method import *
from public import *





def mergeNewValue():
    """合并单位净值(Unit Value) 累计净值(Accumulate Value)"""

    net_value_excel_name = "20221215-固收净值.xlsx"
    unit_value_sheet = pd.read_excel(Public.readFile(net_value_excel_name), sheet_name="单位净值")
    accu_value_sheet = pd.read_excel(Public.readFile(net_value_excel_name), sheet_name="累计净值")



