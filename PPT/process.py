import settings
import common
from sql import price_index_data, price_level_data, manf_price_data
import pandas as pd
import xlwings as xw
from win32com.client import constants as c
from numpy import ceil, floor
import re
import locale
import warnings
warnings.filterwarnings('ignore')
locale.setlocale(locale.LC_CTYPE, 'chinese')
logger = settings.logger


def OpenExcel():
    global app, wb

    app = xw.App(add_book=False, visible=settings.isVisible)
    app.display_alerts = False
    wb = app.books.open(settings.GetRecentFile(settings.sourcePath, '.xlsm'))


def CloseExcel():
    wb.close()
    app.kill()


def GetData(segment, columns):

    data = price_level_data.query("CUST_SEGMENT_NAME_CHN == @segment").reindex(columns, axis=1).fillna(0)
    return data


def GetSummaryData(segment, target, index):

    column = 'TP' if target == '加权成交价汇总' else 'DISCOUNT_RATE'

    if segment == '整体':
        data = price_index_data.loc[:, [column]].reindex(index, axis=0).append(price_index_data.loc[~price_index_data.index.isin(index), [column]])
    elif segment == '对手':
        data = manf_price_data.loc[:, [column]].reindex(index, axis=0).append(manf_price_data.loc[~manf_price_data.index.isin(index), [column]])
    else:
        ws = wb.sheets[target]
        data = ws.range(ws.cells.api.Find(segment).Address).offset(1, 1).options(pd.DataFrame, index=False, header=False, expand='table').value
        data = data.iloc[:, 1:2].applymap(lambda x: str(int(x)) if isinstance(x, (float, int)) else x).join(data.iloc[:, -2:-1]).set_index(1).rename(columns=lambda x: column).append(price_index_data.loc[[segment], [column]])
        data = data.reindex(index, axis=0).append(data[~data.index.isin(index)]).replace('-', '')
    return data.fillna('')


def UpdateShapeChart(data, chart):
    chart.Chart.ChartData.Activate()
    pws = chart.Chart.ChartData.Workbook.Sheets[1]

    if chart.Name == '细分市场价格带':

        max_row = pws.Cells.Columns(1).Cells(pws.Cells.Columns(1).Cells.Count).End(c.xlUp).Row
        max_col = pws.Cells.Rows(1).Cells(pws.Cells.Rows(1).Cells.Count).End(c.xlToLeft).Column

        pws.Range(pws.Cells(max_row + 1, 1), pws.Cells(max_row + 1, max_col)).Value = DataFrameToList(data, index=True)
        pws.Cells(max_row, 1).Value = '{.month}月'.format(pws.Cells(max_row, 1).Value)
        pws.Cells(max_row + 1, 1).Value = '{}年{}月'.format(settings.price_date//10000, settings.price_date%10000//100)

    elif chart.Name in ('细分市场加权均价', '细分市场加权折扣率'):

        max_row = pws.Cells.Columns(1).Cells(pws.Cells.Columns(1).Cells.Count).End(c.xlUp).Row
        max_col = pws.Cells.Rows(1).Cells(pws.Cells.Rows(1).Cells.Count).End(c.xlToLeft).Column

        if not pws.Cells(max_row, 1).Value:
            pws.Cells(max_row, 1).EntireRow.Delete()
            max_row = pws.Cells.Columns(1).Cells(pws.Cells.Columns(1).Cells.Count).End(c.xlUp).Row
        
        pws.Cells(1, max_col).Value = re.sub('\d+年', '   ', pws.Cells(1, max_col).Value)
        pws.Cells(1, max_col + 1).Value = '{}年{}月'.format(settings.price_date//10000, settings.price_date%10000//100)
        
        pws.Range(pws.Cells(2, 1), pws.Cells(data.shape[0] + 1, 1)).Value = DataFrameToList(data.iloc[:, :0], index=True)
        pws.Range(pws.Cells(2, max_col + 1), pws.Cells(data.shape[0] + 1, max_col + 1)).Value = DataFrameToList(data.iloc[:, :1], index=False)
        pws.Range(pws.Cells(2, max_col + 1), pws.Cells(data.shape[0] + 1, max_col + 1)).NumberFormat = pws.Cells(2, max_col).NumberFormat

        AdjustSeriesStyle(chart)
        AdjustChartAxes(chart, pws)

    chart.Chart.DisplayBlanksAs = c.xlNotPlotted
    chart.Chart.Refresh()
    chart.Chart.ChartData.Workbook.Close()
    chart.Chart.Refresh()

    del chart, pws
    return None


def AdjustSeriesStyle(shape):
    for idx in range(1, shape.Chart.SeriesCollection().Count + 1):
        series = shape.Chart.SeriesCollection(idx)
        series.Format.Line.Visible = False
        series.Format.Line.Visible = True
        series.Points(series.Points().Count - 1).HasDataLabel = False
        series.Points(series.Points().Count).HasDataLabel = True
        series.Points(series.Points().Count).DataLabel.Font.Size = 6
        series.Points(series.Points().Count).DataLabel.Text = '{} {}'.format(series.Name, series.Points(series.Points().Count).DataLabel.Text)
        series.Points(series.Points().Count).DataLabel.Font.Color = series.Format.Line.ForeColor.RGB
        series.MarkerStyle = 8                             # 空心圆点
        series.MarkerSize = 5                              # 点大小
        series.Format.Fill.Solid()                         # 纯色填充
        series.MarkerBackgroundColor = RGB(255,255,255)    # 填充白色
        del series
    return None


def AdjustChartAxes(chart, pws):

    max_row = pws.Cells.Columns(1).Cells(pws.Cells.Columns(1).Cells.Count).End(c.xlUp).Row
    max_col = pws.Cells.Rows(1).Cells(pws.Cells.Rows(1).Cells.Count).End(c.xlToLeft).Column

    if chart.Name == '细分市场加权均价':
        max_value = pws.Application.WorksheetFunction.Max(pws.Range(pws.Cells(2, 2), pws.Cells(max_row, max_col)))
        min_value = pws.Application.WorksheetFunction.Min(pws.Range(pws.Cells(2, 2), pws.Cells(max_row, max_col)))
        chart.Chart.Axes(c.xlValue).MaximumScale = int(ceil(max_value / 10000) * 10000)
        chart.Chart.Axes(c.xlValue).MinimumScale = int(floor(min_value / 10000) * 10000)
        chart.Chart.Axes(c.xlValue).MajorUnitIsAuto = True

    elif chart.Name == '细分市场加权折扣率':
        max_value = pws.Application.WorksheetFunction.Max(pws.Range(pws.Cells(2, 2), pws.Cells(max_row, max_col)))
        min_value = pws.Application.WorksheetFunction.Min(pws.Range(pws.Cells(2, 2), pws.Cells(max_row, max_col)))
        chart.Chart.Axes(c.xlValue).MaximumScale = ceil(max_value * 100) / 100
        chart.Chart.Axes(c.xlValue).MinimumScale = 0
        chart.Chart.Axes(c.xlValue).MajorUnitIsAuto = True
    
    del chart, pws
    return None


# 更新 Table
def UpdateShapeTable(data, table, start_row=1, start_column=2):

    if isinstance(data, pd.DataFrame):
        data = DataFrameToList(data)
    
    width = table.Width
    dataRowsCount, dataColsCount,  = len(data), len(data[0])
    tableRowsCount, tableColsCount = table.Table.Rows.Count, table.Table.Columns.Count

    if tableRowsCount < dataRowsCount + start_row - 1:        # shape_table 行不够
        for _ in range(dataRowsCount + start_row - 1 - tableRowsCount):
            table.Table.Rows.Add()
    elif tableRowsCount > dataRowsCount + start_row - 1:      # shape_table 行多余
        for _ in range(tableRowsCount - dataRowsCount - start_row + 1):
            table.Table.Rows(start_row).Delete()
            
    if tableColsCount < dataColsCount + start_column - 1:     # shape_table 列不够
        for _ in range(dataColsCount + start_column - 1 - tableColsCount):
            table.Table.Columns.Add()
    elif tableColsCount > dataColsCount + start_column - 1:   # shape_talbe 列多余
        for _ in range(tableColsCount - dataColsCount - start_column + 1):
            table.Table.Columns(start_column).Delete()

    from itertools import product
    for row, col in product(range(start_row, table.Table.Rows.Count + 1), range(start_column, table.Table.Columns.Count + 1)):
        value = data[row - start_row][col - start_column]
        table.Table.Cell(row, col).Shape.TextFrame.TextRange.Text = value

    table.Width = width
    del table, tableRowsCount, tableColsCount
    return None


'''DataFrame 转为 2dList , 便于输出Excel'''
def DataFrameToList(df, index=False, header=False):
    if header:
        column = [df.columns.tolist()] if df.columns.nlevels == 1 else \
                [list(map(lambda x: x[0], df.columns)), list(map(lambda x: x[1], df.columns))]
        if index:                                              # 是否保留 index 列
            list(map(lambda x: x.insert(0, ''), column))       # 保留 index 需要在表头插入 '' , 只考虑一级 index
            data = column + df.reset_index().values.tolist()
        else:
            data = column + df.values.tolist()
    else:
        data = df.reset_index().values.tolist() if index else df.values.tolist()
        
    return data


def RGB(red, green, blue):
    assert 0 <= red <=255
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)

colorGreen, colorRed = RGB(0, 176, 80), RGB(255, 0, 0)