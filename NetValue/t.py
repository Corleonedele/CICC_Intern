import openpyxl
from Public import Public


wb = openpyxl.load_workbook("净值.xlsx")
st = wb["Sheet1"]

tewb = openpyxl.load_workbook("test.xlsx")
tsst = tewb["Sheet1"]


product_code = Public.readRow(st, 1, 4)
product_name = Public.readRow(st, 2, 4)
date= Public.readColumn(st, 2, 4)
print("数据读取完毕 写入中")
Public.writeRow(tsst, 1, ["净值日期", "产品代码", "产品名称", "单位净值", "累计净值", "除权净值"], 1)

row_index = 4
write_row = 2
for i in date:
    product_info = Public.readRow(st, row_index, 4)
    for ind, value in enumerate(product_info):
        if value == 0: continue
        if value == None:continue   
        # print([i, product_code[ind], product_name[ind], value])
        Public.writeRow(tsst, write_row, [i, product_code[ind], product_name[ind], value, value, value], 1)
        write_row+=1
    row_index+=1
print("净值写入完毕")
tewb.save("test1.xlsx")