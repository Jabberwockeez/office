'''
@Author: Guo Yulong
@Date: 2018-01-15 23:27:04
@LastEditors: Guo Yulong
@LastEditTime: 2019-01-30 19:17:08
'''

import settings
import province_process
import pandas as pd
import xlwings as xw
import win32com.client as com
from win32com.client import constants as c
import re
import os
import warnings
warnings.filterwarnings('ignore')


app = xw.App(add_book=False, visible=settings.isVisible)
app.display_alerts = False
# app.screen_updating = False

template_wb = app.books.open(settings.GetRecentFile(settings.origPath, '省级地图'))
province = template_wb.sheets['省份'].range(1, 1).options(pd.DataFrame, index=False, expand='table').value

source_wb = app.books.open(settings.GetRecentFile(settings.sourcePath, '.xlsx'))
data = source_wb.sheets[0].range(1, 1).options(pd.DataFrame, index=False, expand='table').value.merge(province, left_on='目标城市', right_on='PROVINCE_NAME', how='left', indicator=True)
if 'left_only' in data._merge.values:
    print('不存在省份: ', ' '.join(data.query("_merge == 'left_only'").目标城市.values))
    data = data.query("_merge != 'left_only'")

province_process.UpdateRegion(data, source_wb, template_wb.sheets['地图'])
province_process.UpdateNation(data, source_wb, template_wb.sheets['地图'])

template_wb.close()
# source_wb.save(os.path.join(settings.resultPath, f'MAP_{pd.datetime.now():%Y%m%d%H%M}.xlsx'))
# app.screen_updating = True
app.visible = 1