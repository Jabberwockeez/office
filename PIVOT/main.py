import settings
import process
import xlwings as xw
import pandas as pd
from win32com.client import constants as c
from numpy import nan
import tkinter as tk
import tkinter.simpledialog as simpledialog
import os
import re
import warnings
warnings.filterwarnings('ignore')



def f2(sourceFileName, sheet):

    wb = app.books.open(sourceFileName)
    ws = wb.sheets[sheet]
    data = ws.range(1, 1).options(pd.DataFrame, index=False, expand='table').value.loc[:, :'备注（数据来源）'].fillna('')
    data[['厂商', '品牌']] = data[['厂商', '品牌']].applymap(lambda x: x.replace(',', '、').replace('，', '、'))

    if not '清单（合并）' in (ws.name for ws in wb.sheets):
        ws = wb.sheets.add('清单（合并）', after=wb.sheets.count)
        ws.range(1, 1).options(index=False).value = process.MergeData(data)
    else:
        ws = wb.sheets['清单（合并）']
        ws.range((3, 1), ws.cells.last_cell).clear()
        ws.range(2, 1).options(index=False, header=False).value = process.MergeData(data)

    if not '集团简介' in (ws.name for ws in wb.sheets):
        ws = wb.sheets.add('集团简介', after=wb.sheets.count)
        ws.range(1, 1).options(index=False).value = process.PivotData(data)
    else:
        ws = wb.sheets['集团简介']
        ws.range((3, 1), ws.cells.last_cell).clear()
        ws.range(2, 1).options(index=False, header=False).value = process.PivotData(data)
    return None


def f3(sourceFileName, sheet):

    wb = app.books.open(sourceFileName)
    ws = wb.sheets[sheet]
    data = ws.range(1, 2).options(pd.DataFrame, index=False, expand='table').value.loc[:, :'认缴出资额（万元）'].fillna('')
    data = data[data['出资比例 '].map(str).str.contains('\d', na=False)].drop('工商注册名称', axis=1)

    ws = wb.sheets['汇总']
    ws.range((3, 1), ws.cells.last_cell).clear()
    ws.range(2, 1).options(index=False, header=False).value = process.MergeRankData(data)
    max_row = ws.cells.columns(1).last_cell.end(c.xlUp).row
    max_col = ws.cells.rows(1).last_cell.end(c.xlToLeft).column
    ws.range((2, 1), (2, max_col)).api.Copy()
    ws.range((3, 1), (max_row, max_col)).api.PasteSpecial(c.xlPasteFormats)
    app.api.Application.CutCopyMode = False
    return None


def f5(sourceFileName, sheet):

    wb = app.books.open(sourceFileName)
    ws = wb.sheets[sheet]
    data = ws.range(1, 1).options(pd.DataFrame, index=False, expand='table').value.loc[:, :'原集团'].fillna('')
    data[['厂商', '品牌']] = data[['厂商', '品牌']].applymap(lambda x: x.replace(',', '、').replace('，', '、'))

    if not '清单（合并）' in (ws.name for ws in wb.sheets):
        ws = wb.sheets.add('清单（合并）', after=wb.sheets.count)
        ws.range(1, 1).options(index=False).value = process.MergeData(data)
    else:
        ws = wb.sheets['清单（合并）']
        ws.range((3, 1), ws.cells.last_cell).clear()
        ws.range(2, 1).options(index=False, header=False).value = process.MergeData(data)

    if not '集团简介' in (ws.name for ws in wb.sheets):
        ws = wb.sheets.add('集团简介', after=wb.sheets.count)
        ws.range(1, 1).options(index=False).value = process.PivotData(data)
    else:
        ws = wb.sheets['集团简介']
        ws.range((3, 1), ws.cells.last_cell).clear()
        ws.range(2, 1).options(index=False, header=False).value = process.PivotData(data)
    return None


def f6(sourceFileName, sheet):

    wb = app.books.open(sourceFileName)
    ws = wb.sheets[sheet]
    data = ws.range(1, 1).options(pd.DataFrame, index=False, expand='table').value.loc[:, :'认缴出资额'].fillna('')
    data = data[data['出资比例 '].map(str).str.contains('\d', na=False)]

    ws = wb.sheets['汇总']
    ws.range((3, 1), ws.cells.last_cell).clear()
    ws.range(2, 1).options(index=False, header=False).value = process.MergeRankData(data)
    max_row = ws.cells.columns(1).last_cell.end(c.xlUp).row
    max_col = ws.cells.rows(1).last_cell.end(c.xlToLeft).column
    ws.range((2, 1), (2, max_col)).api.Copy()
    ws.range((3, 1), (max_row, max_col)).api.PasteSpecial(c.xlPasteFormats)
    app.api.Application.CutCopyMode = False
    return None



if __name__ == "__main__":
        
    ROOT = tk.Tk()
    ROOT.withdraw()

    file_dict = {int(re.match('\d+', x).group()):x for x in os.listdir(settings.sourcePath)}
    USER_INPUT = simpledialog.askinteger(title="选择报告", prompt='\n'.join([f'{key} - {value}' for key, value in file_dict.items()]))
    if not USER_INPUT:
        exit(0)
    sourceFileName = settings.GetRecentFile(os.path.join(settings.sourcePath, file_dict.get(USER_INPUT)))

    app = xw.App(add_book=False, visible=settings.isVisible)
    app.display_alerts = False

    if USER_INPUT == 2:
        f2(sourceFileName, '原始数据')
    elif USER_INPUT == 3:
        f3(sourceFileName, '数据源')
    elif USER_INPUT == 5:
        f5(sourceFileName, '非百强集团旗下4S店清单（数据源）')
    elif USER_INPUT == 6:
        f6(sourceFileName, '数据源')

    app.visible = 1
    exit(0)