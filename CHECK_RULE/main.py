import settings
from process import CheckFormat, CheckLogic
from process import colorYellow, colorRed, colorNone
import xlwings as xw
import pandas as pd
from itertools import product
import os
import re
import tkinter as tk
import tkinter.filedialog
import warnings
warnings.filterwarnings('ignore')

logger = settings.logger

app = xw.App(add_book=False, visible=settings.isVisible)
app.display_alerts = False

for reportFileName in os.listdir(settings.sourcePath)[:]:

    sheet_list = pd.ExcelFile(os.path.join(settings.sourcePath, reportFileName)).sheet_names
    if '配置查询(公式引用)' in sheet_list:
        sheet_list.remove('配置查询(公式引用)')
    sheet = '配置查询(去公式)' if '配置查询(去公式)' in sheet_list else sheet_list[0]
    print(f'『{reportFileName}』 :  {sheet}')
    report_data = pd.read_excel(os.path.join(settings.sourcePath, reportFileName), sheet)
    report_data.index = pd.RangeIndex(2, len(report_data)+2)

    format_data = CheckFormat(report_data, reportFileName.split('_')[0])
    mask = format_data.apply(lambda x: any(x.isin([False])), axis=1)
    format_data = format_data[mask]
    f_number = format_data.apply(lambda x: x.value_counts()).sum(axis=1).get(False) if len(format_data) else 0

    logic_data = CheckLogic(report_data)
    mask = logic_data.apply(lambda x: any(x.isin([False])), axis=1)
    logic_data = logic_data[mask]
    l_number = logic_data.apply(lambda x: x.value_counts()).sum(axis=1).get(False) if len(logic_data) else 0

    logger.info('『{}』\n\t\t-- 数据量:{}*{}  格式错误:{}  逻辑错误:{}\n'.format(
        reportFileName, report_data.shape[0], report_data.shape[1]-6, int(f_number), int(l_number)))

    wb = app.books.open(os.path.join(settings.sourcePath, reportFileName))
    if '配置查询(公式引用)' in (ws.name for ws in wb.sheets):
        wb.sheets['配置查询(公式引用)'].delete()
    ws = wb.sheets[sheet]

    for row, col in product(format_data.index, range(format_data.shape[1])):
        if not format_data.loc[row].iloc[col]:
            ws.range(row, col + 7).color = colorYellow

    for row, col in product(logic_data.index, range(logic_data.shape[1])):
        if not logic_data.loc[row].iloc[col]:
            ws.range(row, col + 7).color = colorRed
            
    resultFileName = re.sub('.xlsx', f"_CHECK_{pd.datetime.now().strftime('%Y%m%d%H%M')}.xlsx", reportFileName)
    wb.save(os.path.join(settings.reportPath, resultFileName))
    wb.close()

    # break

app.kill()