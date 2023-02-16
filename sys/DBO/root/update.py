import openpyxl


STANDARD_DATE = "20230203"

class Path():
    交易数据 = "./DBO/DB/UPDATED/私募种子基金持仓日报表.xlsx"
    交易数据_OLD = "./DBO/DB/TEST/私募种子基金持仓日报表.xlsx"



def run():
    modified = ["交易数据"]
    # 交易数据更新
    # 生成持仓日报表
    # 净值更新
    # 生成风险日报表

    if "交易数据" in modified:
        交易数据 = openpyxl.load_workbook(Path.交易数据)["Sheet"]
        交易数据_OLD = openpyxl.load_workbook(Path.交易数据_OLD)["Sheet"]

        print("交易数据更新", 交易数据_OLD.max_row - 交易数据.max_row)
    

    # 清空TEST文件夹
    # 复制UPDATED文件到TEST文件夹

run()