from method import *
from public import *
from WindPy import w




def data():
    # 处理data
    unit_value = Method.mergeSameSheet(file_name="固收净值", sheet_name="单位净值", path="Input_Data", header=[0,1], index_col=[0])
    accu_value = Method.mergeSameSheet(file_name="固收净值", sheet_name="累计净值", path="Input_Data", header=[0,1], index_col=[0])

    history_unit_value = pd.read_excel(Public.readFile("单位净值历史数据.xlsx"))
    history_accu_value = pd.read_excel(Public.readFile("单位净值历史数据.xlsx"))
    code_file = pd.read_excel(Public.readFile("统一三处代码.xlsx"))

    code_SZ = code_file[["金富1号"]].T # 金富1号代码
    unit_value.columns=code_SZ.values[0]
    accu_value.columns=code_SZ.values[0]


    start_date = unit_value.index[0]
    end_date = unit_value.index[-1]
    try:
        w.start("username=W1457909349;password=glh346glh;sitename=NJDX", waitTime=120)
        index_data = w.wsd("000906.SH", "close", start_date, end_date)
        w.close()
    except:
        print("Wind Connect Error")
    

    date_pd = []
    for i in index_data.Times:
        date_pd.append(pd.Timestamp(i))
    index_value = pd.DataFrame(index_data.Data[0], columns=["中证800"], index=date_pd)

    unit_value=pd.concat([unit_value, index_value], ignore_index = False)
    unit_value = unit_value.groupby(unit_value.index).first() 

    accu_value=pd.concat([accu_value, index_value], ignore_index = False)
    accu_value = accu_value.groupby(accu_value.index).first() 

    unit_value = unit_value[unit_value["中证800"] != 0]
    accu_value = accu_value[accu_value["中证800"] != 0]

    unit_value.insert(0, "日期", unit_value.index)
    unit_value.insert(0, "Unnamed: 0", unit_value.index)
    accu_value.insert(0, "日期", accu_value.index)
    accu_value.insert(0, "Unnamed: 0", accu_value.index)


    history_unit_value = pd.concat([history_unit_value, unit_value], ignore_index = True)
    history_accu_value = pd.concat([history_accu_value, accu_value], ignore_index = True)
    history_unit_value = history_unit_value.fillna(0)
    history_accu_value = history_accu_value.fillna(0)
    


    # return history_unit_value, history_accu_value
    history_accu_value.to_csv("Output_Data/mid_1.csv")
    history_unit_value.to_csv("Output_Data/mid_2.csv")
    

data()