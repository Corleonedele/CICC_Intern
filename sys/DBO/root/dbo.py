import os
import openpyxl


DB_PATH = "./DBO/DB/"
REPORT_PATH = "./DBO/DB/REPORT/"

def sheet_copy_paste(to_st, from_st):
    for row in range(1, from_st.max_row+1):
        for col in range(1, from_st.max_column+1):
            to_st.cell(row, col).value = from_st.cell(row, col).value


class 交易记录Method():
    def 公用输入():
        return 0

    def 加权成本(sheet):
        产品清算金额汇总 = 0
        产品持仓数量汇总 = 0

        for row in range(1+1, sheet.max_row+1):
            direction = sheet.cell(row, 2).value
            value = float(sheet.cell(row, 20).value)
            amount = float(sheet.cell(row, 15).value)

            if direction == "分红再投":
                产品清算金额汇总 += value
            elif direction == "现金分红":
                产品清算金额汇总 -= value
            elif direction == "调减":
                产品清算金额汇总 += value
            elif direction == "追加":
                产品清算金额汇总 -= value
            elif direction == "赎回":
                产品清算金额汇总 -= value


            if direction == "分红再投":
                产品持仓数量汇总 += amount
            elif direction == "现金分红":
                产品持仓数量汇总 += amount
            elif direction == "调减":
                产品持仓数量汇总 -= amount
            elif direction == "追加":
                产品持仓数量汇总 += amount
            elif direction == "赎回":
                产品持仓数量汇总 -= amount

        if 产品持仓数量汇总 == 0:
            return 0
        return 产品清算金额汇总 / 产品持仓数量汇总

    def 本年度加权成本(sheet):
        产品本年度成本汇总 = 0
        产品持仓数量变动汇总  = 0

        for row in range(1+1, sheet.max_row):
            direction = sheet.cell(row, 2).value
            value = float(sheet.cell(row, 25).value)
            amount = float(sheet.cell(row, 15).value)

            if direction == "分红再投":
                产品本年度成本汇总 += value
            elif direction == "现金分红":
                产品本年度成本汇总 -= value
            elif direction == "调减":
                产品本年度成本汇总 += value
            elif direction == "追加":
                产品本年度成本汇总 -= value
            elif direction == "赎回":
                产品本年度成本汇总 -= value


            if direction == "分红再投":
                产品持仓数量变动汇总 += amount
            elif direction == "现金分红":
                产品持仓数量变动汇总 += amount
            elif direction == "调减":
                产品持仓数量变动汇总 -= amount
            elif direction == "追加":
                产品持仓数量变动汇总 += amount
            elif direction == "赎回":
                产品持仓数量变动汇总 -= amount

        if 产品持仓数量变动汇总 == 0:
            return 0
        return 产品本年度成本汇总 / 产品持仓数量变动汇总

def init_report():
    sheet_name = ["备注信息表", "私募种子基金持仓日报表", "交易记录", "份额-产品到期日期", "股票多头", "指数增强", "空气指增", "量化择时", "量化对冲", "宏观对冲", "量化期货", "多策略灵活配置", "资金流水", "年初资产+追加资产", "Mapping", "FOF(YTD)"]
    report = openpyxl.Workbook()
    st = report.active
    st.title = "风险控制指标情况"
    for i in sheet_name:
        report.create_sheet(i)

    info = openpyxl.load_workbook(DB_PATH+"公共信息.xlsx")

    sheet_copy_paste(report["备注信息表"], info["备注信息表"])
    sheet_copy_paste(report["份额-产品到期日期"], info["份额-产品到期日期"])
    sheet_copy_paste(report["资金流水"], info["资金流水"])
    sheet_copy_paste(report["年初资产+追加资产"], info["年初资产+追加资产"])
    sheet_copy_paste(report["Mapping"], info["Mapping"])


    
    report.save(REPORT_PATH+"report.xlsx")



def 追加(body_dict):
    wb = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    st = wb["Sheet1"]
    write_row = st.max_row + 1


    # 手动输入
    # 成交时间 col = 1
    st.cell(write_row, 1).value = body_dict["成交时间"]
    # 买卖方向 col = 2
    st.cell(write_row, 2).value = "追加"
    # 证劵代码 col = 3
    st.cell(write_row, 3).value = body_dict["证劵代码"]
    # 产品名称 col = 4
    st.cell(write_row, 4).value = body_dict["产品名称"]
    # 产品管理人 col = 5
    st.cell(write_row, 5).value = body_dict["产品管理人"]
    # 策略类型 col = 6
    st.cell(write_row, 6).value = body_dict["策略类型"]
    # 策略类型_新 col = 7
    st.cell(write_row, 7).value = body_dict["策略类型_新"]
    # 跟踪指数 col = 8
    st.cell(write_row, 8).value = body_dict["跟踪指数"]
    # 细分策略 col = 9
    st.cell(write_row, 9).value = body_dict["细分策略"]
    # 产品分类 col = 10
    st.cell(write_row, 10).value = body_dict["产品分类"]
    # 初始投资金额 col = 14
    st.cell(write_row, 14).value = body_dict["初始投资金额"]
    # 成交数量 col = 15
    st.cell(write_row, 15).value = body_dict["成交数量"]
    # 成交金额_万元 col = 16
    st.cell(write_row, 16).value = body_dict["成交金额_万元"]
    # 本年度成本价 col = 24
    st.cell(write_row, 24).value = body_dict["本年度成本价"]
    # 分支机构 col =11
    st.cell(write_row, 11).value = body_dict["分支机构"]
    # 推荐IC col = 12
    st.cell(write_row, 12).value = body_dict["推荐IC"]
    # 考核承担IC col = 13
    st.cell(write_row, 13).value = body_dict["考核承担IC"]
    # IC分摊比例 = 17
    st.cell(write_row, 17).value = body_dict["IC分摊比例"]


    # 公式计算

    # 初始成交价 col = 18
    st.cell(write_row, 18).value = float(body_dict["成交金额_万元"]) / float(body_dict["成交数量"])
    # 数量变动 col = 19
    st.cell(write_row, 19).value = float(body_dict["成交数量"])
    # 清算金额 col = 20
    st.cell(write_row, 20).value = (-1) * float(body_dict["成交金额_万元"])
    # 持仓金额变动 col = 21
    st.cell(write_row, 21).value = (-1) * float(body_dict["成交金额_万元"])
    # 加权成本 col = 22
    st.cell(write_row, 22).value = 交易记录Method.加权成本(st)
    # 卖出成本 col = 23
    st.cell(write_row, 23).value = 0
    # 本年度成本 col = 25
    st.cell(write_row, 25).value = 0
    # 本年度加权成本 col = 26
    st.cell(write_row, 26).value = 交易记录Method.本年度加权成本(st)
    # 本年度卖出成本 col = 27
    st.cell(write_row, 27).value = 0
    # 累计已实现收益 col = 28
    st.cell(write_row, 28).value = 0
    # 本年度已实现收益 col = 29
    st.cell(write_row, 29).value = 0

    wb.save(DB_PATH+"交易记录.xlsx")


def 调减(body_dict):
    wb = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    st = wb["Sheet1"]
    write_row = st.max_row + 1

    # 手动输入
    # 成交时间 col = 1
    st.cell(write_row, 1).value = body_dict["成交时间"]
    # 买卖方向 col = 2
    st.cell(write_row, 2).value = "调减"
    # 证劵代码 col = 3
    st.cell(write_row, 3).value = body_dict["证劵代码"]
    # 产品名称 col = 4
    st.cell(write_row, 4).value = body_dict["产品名称"]
    # 产品管理人 col = 5
    st.cell(write_row, 5).value = body_dict["产品管理人"]
    # 策略类型 col = 6
    st.cell(write_row, 6).value = body_dict["策略类型"]
    # 策略类型_新 col = 7
    st.cell(write_row, 7).value = body_dict["策略类型_新"]
    # 跟踪指数 col = 8
    st.cell(write_row, 8).value = body_dict["跟踪指数"]
    # 细分策略 col = 9
    st.cell(write_row, 9).value = body_dict["细分策略"]
    # 产品分类 col = 10
    st.cell(write_row, 10).value = body_dict["产品分类"]
    # 初始投资金额 col = 14
    st.cell(write_row, 14).value = body_dict["初始投资金额"]
    # 成交数量 col = 15
    st.cell(write_row, 15).value = body_dict["成交数量"]
    # 成交金额_万元 col = 16
    st.cell(write_row, 16).value = body_dict["成交金额_万元"]
    # 本年度成本价 col = 24
    st.cell(write_row, 24).value = body_dict["本年度成本价"]
    # 分支机构 col =11
    st.cell(write_row, 11).value = body_dict["分支机构"]
    # 推荐IC col = 12
    st.cell(write_row, 12).value = body_dict["推荐IC"]
    # 考核承担IC col = 13
    st.cell(write_row, 13).value = body_dict["考核承担IC"]
    # IC分摊比例 = 17
    st.cell(write_row, 17).value = body_dict["IC分摊比例"]


    # 公式计算

    # 初始成交价 col = 18
    st.cell(write_row, 18).value = float(body_dict["成交金额_万元"]) / float(body_dict["成交数量"])
    # 数量变动 col = 19
    st.cell(write_row, 19).value = float(body_dict["成交数量"])
    # 清算金额 col = 20
    st.cell(write_row, 20).value = (-1) * float(body_dict["成交金额_万元"])
    # 持仓金额变动 col = 21
    st.cell(write_row, 21).value = (-1) * float(body_dict["成交金额_万元"])
    # 加权成本 col = 22
    st.cell(write_row, 22).value = 交易记录Method.加权成本(st)
    # 卖出成本 col = 23
    st.cell(write_row, 23).value = 0
    # 本年度成本 col = 25
    st.cell(write_row, 25).value = 0
    # 本年度加权成本 col = 26
    st.cell(write_row, 26).value = 交易记录Method.本年度加权成本(st)
    # 本年度卖出成本 col = 27
    st.cell(write_row, 27).value = 0
    # 累计已实现收益 col = 28
    st.cell(write_row, 28).value = 0
    # 本年度已实现收益 col = 29
    st.cell(write_row, 29).value = 0


    wb.save(DB_PATH+"交易记录.xlsx")


def 赎回(body_dict):
    wb = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    st = wb["Sheet1"]
    write_row = st.max_row + 1

    # 手动输入
    # 成交时间 col = 1
    st.cell(write_row, 1).value = body_dict["成交时间"]
    # 买卖方向 col = 2
    st.cell(write_row, 2).value = "赎回"
    # 证劵代码 col = 3
    st.cell(write_row, 3).value = body_dict["证劵代码"]
    # 产品名称 col = 4
    st.cell(write_row, 4).value = body_dict["产品名称"]
    # 产品管理人 col = 5
    st.cell(write_row, 5).value = body_dict["产品管理人"]
    # 策略类型 col = 6
    st.cell(write_row, 6).value = body_dict["策略类型"]
    # 策略类型_新 col = 7
    st.cell(write_row, 7).value = body_dict["策略类型_新"]
    # 跟踪指数 col = 8
    st.cell(write_row, 8).value = body_dict["跟踪指数"]
    # 细分策略 col = 9
    st.cell(write_row, 9).value = body_dict["细分策略"]
    # 产品分类 col = 10
    st.cell(write_row, 10).value = body_dict["产品分类"]
    # 初始投资金额 col = 14
    st.cell(write_row, 14).value = body_dict["初始投资金额"]
    # 成交数量 col = 15
    st.cell(write_row, 15).value = body_dict["成交数量"]
    # 成交金额_万元 col = 16
    st.cell(write_row, 16).value = body_dict["成交金额_万元"]
    # 本年度成本价 col = 24
    st.cell(write_row, 24).value = body_dict["本年度成本价"]
    # 分支机构 col =11
    st.cell(write_row, 11).value = body_dict["分支机构"]
    # 推荐IC col = 12
    st.cell(write_row, 12).value = body_dict["推荐IC"]
    # 考核承担IC col = 13
    st.cell(write_row, 13).value = body_dict["考核承担IC"]
    # IC分摊比例 = 17
    st.cell(write_row, 17).value = body_dict["IC分摊比例"]


    # 公式计算

    # 初始成交价 col = 18
    st.cell(write_row, 18).value = float(body_dict["成交金额_万元"]) / float(body_dict["成交数量"])
    # 数量变动 col = 19
    st.cell(write_row, 19).value = float(body_dict["成交数量"])
    # 清算金额 col = 20
    st.cell(write_row, 20).value = (-1) * float(body_dict["成交金额_万元"])
    # 持仓金额变动 col = 21
    st.cell(write_row, 21).value = (-1) * float(body_dict["成交金额_万元"])
    # 加权成本 col = 22
    st.cell(write_row, 22).value = 交易记录Method.加权成本(st)
    # 卖出成本 col = 23
    st.cell(write_row, 23).value = 0
    # 本年度成本 col = 25
    st.cell(write_row, 25).value = 0
    # 本年度加权成本 col = 26
    st.cell(write_row, 26).value = 交易记录Method.本年度加权成本(st)
    # 本年度卖出成本 col = 27
    st.cell(write_row, 27).value = 0
    # 累计已实现收益 col = 28
    st.cell(write_row, 28).value = 0
    # 本年度已实现收益 col = 29
    st.cell(write_row, 29).value = 0


    wb.save(DB_PATH+"交易记录.xlsx")


def 现金分红(body_dict):
    wb = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    st = wb["Sheet1"]
    write_row = st.max_row + 1

    # 手动输入
    # 成交时间 col = 1
    st.cell(write_row, 1).value = body_dict["成交时间"]
    # 买卖方向 col = 2
    st.cell(write_row, 2).value = "现金分红"
    # 证劵代码 col = 3
    st.cell(write_row, 3).value = body_dict["证劵代码"]
    # 产品名称 col = 4
    st.cell(write_row, 4).value = body_dict["产品名称"]
    # 产品管理人 col = 5
    st.cell(write_row, 5).value = body_dict["产品管理人"]
    # 策略类型 col = 6
    st.cell(write_row, 6).value = body_dict["策略类型"]
    # 策略类型_新 col = 7
    st.cell(write_row, 7).value = body_dict["策略类型_新"]
    # 跟踪指数 col = 8
    st.cell(write_row, 8).value = body_dict["跟踪指数"]
    # 细分策略 col = 9
    st.cell(write_row, 9).value = body_dict["细分策略"]
    # 产品分类 col = 10
    st.cell(write_row, 10).value = body_dict["产品分类"]
    # 初始投资金额 col = 14
    st.cell(write_row, 14).value = body_dict["初始投资金额"]
    # 成交数量 col = 15
    st.cell(write_row, 15).value = body_dict["成交数量"]
    # 成交金额_万元 col = 16
    st.cell(write_row, 16).value = body_dict["成交金额_万元"]
    # 本年度成本价 col = 24
    st.cell(write_row, 24).value = body_dict["本年度成本价"]
    # 分支机构 col =11
    st.cell(write_row, 11).value = body_dict["分支机构"]
    # 推荐IC col = 12
    st.cell(write_row, 12).value = body_dict["推荐IC"]
    # 考核承担IC col = 13
    st.cell(write_row, 13).value = body_dict["考核承担IC"]
    # IC分摊比例 = 17
    st.cell(write_row, 17).value = body_dict["IC分摊比例"]


    # 公式计算

    # 初始成交价 col = 18
    st.cell(write_row, 18).value = float(body_dict["成交金额_万元"]) / float(body_dict["成交数量"])
    # 数量变动 col = 19
    st.cell(write_row, 19).value = float(body_dict["成交数量"])
    # 清算金额 col = 20
    st.cell(write_row, 20).value = (-1) * float(body_dict["成交金额_万元"])
    # 持仓金额变动 col = 21
    st.cell(write_row, 21).value = (-1) * float(body_dict["成交金额_万元"])
    # 加权成本 col = 22
    st.cell(write_row, 22).value = 交易记录Method.加权成本(st)
    # 卖出成本 col = 23
    st.cell(write_row, 23).value = 0
    # 本年度成本 col = 25
    st.cell(write_row, 25).value = 0
    # 本年度加权成本 col = 26
    st.cell(write_row, 26).value = 交易记录Method.本年度加权成本(st)
    # 本年度卖出成本 col = 27
    st.cell(write_row, 27).value = 0
    # 累计已实现收益 col = 28
    st.cell(write_row, 28).value = 0
    # 本年度已实现收益 col = 29
    st.cell(write_row, 29).value = 0


    wb.save(DB_PATH+"交易记录.xlsx")

def 分红再投(body_dict):
    wb = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    st = wb["Sheet1"]
    write_row = st.max_row + 1

    # 手动输入
    # 成交时间 col = 1
    st.cell(write_row, 1).value = body_dict["成交时间"]
    # 买卖方向 col = 2
    st.cell(write_row, 2).value = "分红再投"
    # 证劵代码 col = 3
    st.cell(write_row, 3).value = body_dict["证劵代码"]
    # 产品名称 col = 4
    st.cell(write_row, 4).value = body_dict["产品名称"]
    # 产品管理人 col = 5
    st.cell(write_row, 5).value = body_dict["产品管理人"]
    # 策略类型 col = 6
    st.cell(write_row, 6).value = body_dict["策略类型"]
    # 策略类型_新 col = 7
    st.cell(write_row, 7).value = body_dict["策略类型_新"]
    # 跟踪指数 col = 8
    st.cell(write_row, 8).value = body_dict["跟踪指数"]
    # 细分策略 col = 9
    st.cell(write_row, 9).value = body_dict["细分策略"]
    # 产品分类 col = 10
    st.cell(write_row, 10).value = body_dict["产品分类"]
    # 初始投资金额 col = 14
    st.cell(write_row, 14).value = body_dict["初始投资金额"]
    # 成交数量 col = 15
    st.cell(write_row, 15).value = body_dict["成交数量"]
    # 成交金额_万元 col = 16
    st.cell(write_row, 16).value = body_dict["成交金额_万元"]
    # 本年度成本价 col = 24
    st.cell(write_row, 24).value = body_dict["本年度成本价"]
    # 分支机构 col =11
    st.cell(write_row, 11).value = body_dict["分支机构"]
    # 推荐IC col = 12
    st.cell(write_row, 12).value = body_dict["推荐IC"]
    # 考核承担IC col = 13
    st.cell(write_row, 13).value = body_dict["考核承担IC"]
    # IC分摊比例 = 17
    st.cell(write_row, 17).value = body_dict["IC分摊比例"]


    # 公式计算

    # 初始成交价 col = 18
    st.cell(write_row, 18).value = float(body_dict["成交金额_万元"]) / float(body_dict["成交数量"])
    # 数量变动 col = 19
    st.cell(write_row, 19).value = float(body_dict["成交数量"])
    # 清算金额 col = 20
    st.cell(write_row, 20).value = (-1) * float(body_dict["成交金额_万元"])
    # 持仓金额变动 col = 21
    st.cell(write_row, 21).value = (-1) * float(body_dict["成交金额_万元"])
    # 加权成本 col = 22
    st.cell(write_row, 22).value = 交易记录Method.加权成本(st)
    # 卖出成本 col = 23
    st.cell(write_row, 23).value = 0
    # 本年度成本 col = 25
    st.cell(write_row, 25).value = 0
    # 本年度加权成本 col = 26
    st.cell(write_row, 26).value = 交易记录Method.本年度加权成本(st)
    # 本年度卖出成本 col = 27
    st.cell(write_row, 27).value = 0
    # 累计已实现收益 col = 28
    st.cell(write_row, 28).value = 0
    # 本年度已实现收益 col = 29
    st.cell(write_row, 29).value = 0


    wb.save(DB_PATH+"交易记录.xlsx")




def 底层资产私募配置情况():
    template = openpyxl.load_workbook(DB_PATH+"持仓日报模版.xlsx")
    trade_record = openpyxl.load_workbook(DB_PATH+"交易记录.xlsx")
    report = openpyxl.load_workbook(REPORT_PATH+"report.xlsx")

    # 写入模版

    rep_st = report["私募种子基金持仓日报表"]
    tem_st = template["底层资产模板"]

    for col in range(1, tem_st.max_column):
        rep_st.cell(1, col).value = tem_st.cell(1, col).value

    for row in range(2, tem_st.max_row):
        rep_st.cell(row, 1).value = tem_st.cell(row, 1).value
        rep_st.cell(row, 2).value = tem_st.cell(row, 2).value
        rep_st.cell(row, 3).value = tem_st.cell(row, 3).value
        rep_st.cell(row, 4).value = tem_st.cell(row, 4).value
        rep_st.cell(row, 5).value = tem_st.cell(row, 5).value
        rep_st.cell(row, 6).value = tem_st.cell(row, 6).value
        rep_st.cell(row, 7).value = tem_st.cell(row, 7).value
        rep_st.cell(row, 8).value = tem_st.cell(row, 8).value


    # 写入数据并计算
    trade_code = []
    tra_st = trade_record["Sheet1"]
    for row in range(1+1, tra_st.max_row+1):
        trade_code.append(tra_st.cell(row, 3).value)
    
    




    report.save(REPORT_PATH+"report.xlsx")



# 追加({'成交时间':'123','买卖方向': '123', '证劵代码': '123', '产品名称': '123', '产品管理人': '12', '策略类型': '1', '策略类型_新': '2', '跟踪指数': '123', '细分策略': '123', '产品分类': '123', '初始投资金额': '123', '成交数量': '123', '成交金额_万元': '123', '本年度成本价': '123', '分支机构': '123', '推荐IC': '1231', '考核承担IC': '23', 'IC分摊比例': '123'})
# 底层资产私募配置情况()
init_report()