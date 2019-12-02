import settings
from process import TransformData, AdjustSheetFormat
from win32com.client import constants as c
import pandas as pd
import xlwings as xw
import os
import warnings
warnings.filterwarnings('ignore')


sourceFileName = settings.GetRecentFile(settings.sourcePath)

app = xw.App(add_book=False, visible=settings.isVisible)
app.display_alerts = False

wb = app.books.open(sourceFileName)
data = wb.sheets[0].range(1, 1).options(pd.DataFrame, index=False, expand='table').value.query(f'YM_ID >= {settings.last_report_date//100}').query('ROLL_AVG_SALES > 0').reset_index(drop=True)

# 筛选两个月都有TP的数据 并重新归一mix
filter_index = data[['VERSION_ID', 'YM_ID', 'TP']].set_index(['VERSION_ID', 'YM_ID']).unstack('YM_ID').apply(lambda x: all(pd.notnull(x)), axis=1)
data = data[data.VERSION_ID.isin(filter_index[filter_index].index.values)].reset_index(drop=True)
data.ROLL_MIX = data.groupby(['DIMENSION', 'YM_ID'])['ROLL_AVG_SALES'].apply(lambda x: x / x.sum())

page_mask_dict = {
    'A00_HB_BEV': (data.SEGMENT == 'A00') & (data.BODY_TYPE == 'HB') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'Conventional_BEV': (data.FUEL_TYPE == 'BEV') & (data['VGC-MARKET'] == 'Conventional'),
    'A_NB_BEV': (data.SEGMENT == 'A') & (data.BODY_TYPE == 'NB') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'A_HB_BEV': (data.SEGMENT == 'A') & (data.BODY_TYPE == 'HB') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'A_SUV_BEV': (data.SEGMENT == 'A') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'A0_SUV_BEV': (data.SEGMENT == 'A0') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'B_SUV_BEV': (data.SEGMENT == 'B') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Volume') ,
    'Premium_BEV': (data.FUEL_TYPE == 'BEV') & (data.PROPERTY == 'Premium') & (data.BRAND == '特斯拉'),
    'PHEV_total': (data.FUEL_TYPE == 'PHEV'),
    'A_NB_PHEV': (data.SEGMENT == 'A') & (data.BODY_TYPE == 'NB') & (data.FUEL_TYPE == 'PHEV') & (data.PROPERTY == 'Volume') ,
    'A0_SUV_PHEV': (data.SEGMENT == 'A0') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'PHEV') & (data.PROPERTY == 'Volume') ,
    'A_SUV_PHEV': (data.SEGMENT == 'A') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'PHEV') & (data.PROPERTY == 'Volume') ,
    'B_SUV_PHEV': (data.SEGMENT == 'B') & (data.BODY_TYPE == 'SUV') & (data.FUEL_TYPE == 'PHEV') & (data.PROPERTY == 'Volume') ,
    'Premium_PHEV': (data.FUEL_TYPE == 'PHEV') & (data.PROPERTY == 'Premium') & (data.BRAND.isin(('奥迪', '宝马', '路虎', '保时捷', '沃尔沃'))),
}

for page, mask in page_mask_dict.items():
    ws = wb.sheets.add(page, after=wb.sheets.count)
    ws.range(1, 1).options(index=True).value = TransformData(data[mask].reset_index())
    AdjustSheetFormat(ws)


wb.save(os.path.join(settings.reportPath, f'PPT_Base_{settings.report_date//100}_{pd.datetime.now():%Y%m%d%H%M}'))
wb.close()
app.kill()