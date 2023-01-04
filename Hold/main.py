from method import *
from public import *
from WindPy import w

INDEX_RATE = {}
INDEX_MAXDRAWDOWN = {}

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
    

def thresholdCal(dataFrame, threshold_desc, output_threshold, df_底层资产私募配置情况) -> pd.DataFrame:
    """计算每类产品的指数, 返回DataFrame"""
    result_df = pd.DataFrame(threshold_desc)

    date = dataFrame.iloc[:,[0]]
    start_date = date.iloc[2]
    end_data = date.iloc[-1]
    index = dataFrame.iloc[:,[1]]
    accu_index_rate = float(index.iloc[2]) / float(index.iloc[-1]) - 1

    for i in range(2, dataFrame.shape[1]):
        product = dataFrame.iloc[:,[i]]
        product_name = product.iloc[0]
        product_value = product.iloc[2:dataFrame.shape[0]]
        # 对每个产品计算

        # 产品累计盈亏指标
        accu_profit_rate = float(df_底层资产私募配置情况.loc[df_底层资产私募配置情况["产品代码"] == dataFrame.columns[i]]["累计盈亏比例"].values[0][:-1]) / 100 #累计亏损指标
        # 指数最新回撤（第一天到最后一天）
        index_rate = float(INDEX_RATE["中证500"])
        # 指数最大回撤（max那天到最后一天）
        index_maxDrawdown = float(INDEX_MAXDRAWDOWN["中证500"])
        # 指数对比相对亏损
        compare_profit_rate = accu_profit_rate - index_rate 
        # 产品最新最大回撤
        try:
            new_maxDrawdown = (float(max(product_value.values)) - float(product_value.values[-1])) / float(max(product_value.values))
        except:
            print(product_name.values, "Error")
            new_maxDrawdown = 0
        # 产品历史最大回撤
        tem = []
        for value in product_value.values:
            t = float(value)
            if t == 0: continue
            tem.append(t)
        # history_maxDrawdown = Compute.maxDrawdown(tem)[0]

        # 产品持有期间盈亏
        try:
            accu_hold_profit_rate = tem[0] / tem[-1] - 1
        except:
            accu_hold_profit_rate = 0
        # 产品对比相对回撤
        compare_maxDrawdown_rate = new_maxDrawdown - index_maxDrawdown


        # 写入DataFrame
        result_df.insert(i, product_name.to_string(), 0)
        # print(accu_profit_rate, index_rate, compare_profit_rate, new_maxDrawdown, )
        if ("产品累计亏损" in output_threshold):
            result_df.iloc[output_threshold.index("产品累计亏损"), i] = accu_profit_rate

        if ("产品相对亏损_不比较" in output_threshold):
            result_df.iloc[output_threshold.index("产品相对亏损_不比较"), i] = compare_profit_rate
        elif ("产品相对亏损" in output_threshold):
            if index_rate < 0.1:
                result_df.iloc[output_threshold.index("产品相对亏损"), i] = compare_profit_rate
            else:
                result_df.iloc[output_threshold.index("产品相对亏损")+1, i] = compare_profit_rate

        if ("产品最新回撤" in output_threshold):
            result_df.iloc[output_threshold.index("产品最新回撤"), i] = new_maxDrawdown
        # if ("产品最大回撤" in output_threshold):
        #     result_df.iloc[output_threshold.index("产品最大回撤" ), i] = history_maxDrawdown
        if ("产品相对回撤" in output_threshold):
            result_df.iloc[output_threshold.index("产品相对回撤"), i] = compare_maxDrawdown_rate
        if ("持有期间跌幅" in output_threshold):
            result_df.iloc[output_threshold.index("持有期间跌幅"), i] = accu_hold_profit_rate

        if ("日期" in output_threshold):
            result_df.iloc[output_threshold.index("日期"), i] = date.iloc[-1]

    return result_df




data()