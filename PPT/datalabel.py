import settings
import pandas as pd
import xlwings as xw
import win32com.client as com
from win32com.client import constants as c
import os
import re
import timeit



def RGB(red, green, blue):
    assert 0 <= red <=255
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)

colorGreen, colorRed = RGB(0, 176, 80), RGB(255, 0, 0)


def UpdateSeriesStyle(shape):
    for idx in range(1, shape.Chart.SeriesCollection().Count + 1):
        series = shape.Chart.SeriesCollection(idx)
        series.Format.Line.Visible = False
        series.Format.Line.Visible = True
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



reportFileName = settings.GetRecentFile(settings.reportPath, '价格分析报告')
App = com.gencache.EnsureDispatch('Powerpoint.Application')
ppt = App.Presentations.Open(reportFileName, WithWindow=settings.isVisible)
App.WindowState = 2    # 窗口最小化

slide = ppt.Slides(31)
# slide.Select()
shape = slide.Shapes('Object 2')
begin = timeit.default_timer()
UpdateSeriesStyle(shape)
print(timeit.default_timer() - begin)