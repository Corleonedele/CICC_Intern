import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, Protection

REPORT_PATH = "./DBO/DB/REPORT/test_report.xlsx"



wb = openpyxl.load_workbook(REPORT_PATH)


def 风险控制指标情况(wb):
    st = wb["风险控制指标情况"]
    print(st.max_row)

风险控制指标情况(wb)