'''
@Author: Guo Yulong
@Date: 2019-01-06 15:44:54
@LastEditors: Guo Yulong
@LastEditTime: 2019-01-30 19:37:54
'''

import pandas as pd
from win32com.client import constants as c


font_size = 16
font_name = '思源黑体 CN Light'


def RGB(red, green, blue): 
    assert 0 <= red <= 255
    assert 0 <= green <= 255
    assert 0 <= blue <= 255
    return red + (green << 8) + (blue << 16)

colorDict = {
    1: RGB(198, 217, 248), 
    2: RGB(85, 142, 213), 
    3: RGB(31, 78, 121), 
    4: RGB(23, 55, 94)
    }
colorNone, colorBlack, colorGray = RGB(255, 255, 255), RGB(0, 0, 0), RGB(240, 240, 240)


def UpdateRegion(data, source_wb, map_ws):

    for region in data.REGION_NAME.unique():
        map_ws.api.Copy(After=source_wb.sheets[source_wb.sheets.count - 1].api)
        ws = source_wb.sheets[source_wb.sheets.count - 1]
        ws.name = region

        ws_data = data.query('REGION_NAME == @region')
        for item in ws.shapes('组合 1').impl.api.GroupItems:

            if item.Type == 5:
                if item.Name not in ws_data.PROVINCE_ABBR:
                    item.Delete()
                else:
                    if item.name not in ws_data[pd.notnull(ws_data.目标城市)].PROVINCE_ABBR.values:
                        item.Fill.ForeColor.RGB = colorGray
                    else:
                        item.Fill.ForeColor.RGB = colorDict[ws_data.query('PROVINCE_ABBR == @item.name')['颜色'].values[0]]
                    # item.Fill.ForeColor.RGB = colorDict[ws_data.query('PROVINCE_ABBR == @item.Name')['颜色'].values[0]]

            elif item.Type == 17:
                if item.TextFrame2.TextRange.Text not in ws_data.PROVINCE_ABBR:
                    item.Delete()
                else:
                    # for idx, ct in enumerate(ws_data.query('PROVINCE_ABBR == @item.TextFrame2.TextRange.Text').CITY_ABBR.values):
                    #     city_text = item.Duplicate()
                    #     city_text.TextFrame2.TextRange.Text = ct
                    #     city_text.Left, city_text.Top = item.Left+5*idx, item.Top+5*idx
                    # item.Delete()
                    item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorBlack
                    item.TextFrame.Characters().Font.Size = font_size
                    item.TextFrame.Characters().Font.Name = font_name
                    item.TextFrame2.TextRange.Font.NameFarEast = font_name
                    # item.TextFrame2.TextRange.Text = ws_data.query('MAP_ENAME == @item.TextFrame2.TextRange.Text')._ABBR.values[0]

    return None


def UpdateNation(data, source_wb, map_ws):

    map_ws.api.Copy(After=source_wb.sheets[source_wb.sheets.count - 1].api)
    ws = source_wb.sheets[source_wb.sheets.count - 1]
    ws.name = '全国'

    for item in ws.shapes('组合 1').impl.api.GroupItems:

        if item.Type == 5:
            item.Fill.ForeColor.RGB = colorNone

        elif item.Type == 17:
            if item.TextFrame2.TextRange.Text not in data.PROVINCE_ABBR.unique():
                item.Delete()
            else:
                # for idx, ct in enumerate(data.query('PROVINCE_ABBR == @item.TextFrame2.TextRange.Text').CITY_ABBR.values):
                #     city_text = item.Duplicate()
                #     city_text.TextFrame2.TextRange.Text = ct
                #     city_text.Left, city_text.Top = item.Left+5*idx, item.Top+5*idx
                # item.Delete()
                item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorBlack
                item.TextFrame.Characters().Font.Size = font_size
                item.TextFrame.Characters().Font.Name = font_name
                item.TextFrame2.TextRange.Font.NameFarEast = font_name
                # item.TextFrame2.TextRange.Text = data.query('MAP_ENAME == @item.TextFrame2.TextRange.Text').CITY_ABBR.values[0]
    return None