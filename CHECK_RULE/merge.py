import settings
import xlwings as xw
import pandas as pd
from itertools import product
from win32com.client import constants as c
import os
import shutil
import re
import tkinter as tk
import tkinter.filedialog
import warnings
warnings.filterwarnings('ignore')


app = xw.App(add_book=False, visible=settings.isVisible)
app.display_alerts = False

for idx, report_file in enumerate((file for file in os.listdir(settings.reportPath) if file.endswith('.xls'))):

    print(report_file)
    if idx == 0:
        head_file = report_file
        wb = app.books.open(os.path.join(settings.reportPath, report_file))
        if '配置查询(公式引用)' in wb.sheets:
            wb.sheets['配置查询(公式引用)'].delete()
        sheet = '配置查询(去公式)' if '配置查询(去公式)' in wb.sheets else wb.sheets[0]
        ws = wb.sheets[sheet]
        last_col = ws.cells.rows(2).last_cell.end(c.xlToLeft).column
        ws.range((1, 7), (1, last_col)).value = report_file
    else:
        tmp_wb = app.books.open(os.path.join(settings.reportPath, report_file))
        if '配置查询(公式引用)' in (tmp_ws.name for tmp_ws in tmp_wb.sheets):
            tmp_wb.sheets['配置查询(公式引用)'].delete()
        sheet = '配置查询(去公式)' if '配置查询(去公式)' in tmp_wb.sheets else tmp_wb.sheets[0]
        tmp_ws = tmp_wb.sheets[sheet]
        last_row = tmp_ws.cells.columns(1).last_cell.end(c.xlUp).row
        last_col = tmp_ws.cells.rows(2).last_cell.end(c.xlToLeft).column
        tmp_ws.range((1, 7), (1, last_col)).value = report_file
        if (last_col - 7 + 1) >= (ws.cells.rows(2).last_cell.column - ws.cells.rows(2).last_cell.end(c.xlToLeft).column):
            tmp_wb.close()
            break
        tmp_ws.range((1, 7), (last_row, last_col)).api.Copy()
        ws.cells.rows(2).last_cell.end(c.xlToLeft).offset(-1, 1).api.PasteSpecial(c.xlPasteAll)
        app.api.Application.CutCopyMode = False
        app.api.Application.SelectionMode = False
        tmp_wb.close()
        # if not os.path.isdir(os.path.join(settings.reportPath, 'Finish')):
        #     os.makedirs()
        del tmp_wb
        shutil.move(os.path.join(settings.reportPath, report_file), os.path.join(settings.finishPath, report_file))


ws.range((1, 7), ws.used_range.last_cell).column_width = 20
wb.save('Report_{:%Y%m%d%H%M}.xls'.format(pd.datetime.now()))
wb.close()
app.kill()
shutil.move(os.path.join(settings.reportPath, head_file), os.path.join(settings.finishPath, head_file))
