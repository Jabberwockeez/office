import settings
import pandas as pd
import math
from win32com.client import constants as c
import os
import xlwings as xw


def TransformData(data):
    from sql import brand_data

    data.BRAND = data.merge(brand_data, on='VERSION_ID', how='left')['BRAND_ENAME']

    last_data = data.query(f'YM_ID == {settings.last_report_date//100}').groupby(data.apply(lambda x: f'{x.BRAND} {x.MODEL_NAME}', axis=1))\
        .apply(lambda x: pd.Series({
            'DISCOUNTR_MOM': (x.DISCOUNTR * x.ROLL_MIX).sum(),
        }))
    curr_data = data.query(f'YM_ID == {settings.report_date//100}').groupby(data.apply(lambda x: f'{x.BRAND} {x.MODEL_NAME}', axis=1))\
        .apply(lambda x: pd.Series({
            'MSRP_MIN': '{:,.1f}'.format(normal_round(x.MSRP.min(), -2)/1000).replace('.', ','),
            '-': '-',
            'MSRP_MAX': '{:,.1f}'.format(normal_round(x.MSRP.max(), -2)/1000).replace('.', ','),
            'CP_MIN': '{:,.1f}'.format(normal_round(x.CP.min(), -2)/1000).replace('.', ','),
            '- ': '-',
            'CP_MAX': '{:,.1f}'.format(normal_round(x.CP.max(), -2)/1000).replace('.', ','),
            'TP_MIN': '{:,.1f}'.format(normal_round(x.TP.min(), -2)/1000).replace('.', ','),
            '-  ': '-',
            'TP_MAX': '{:,.1f}'.format(normal_round(x.TP.max(), -2)/1000).replace('.', ','),
            'E-Range': '/'.join(x['E-RANGE'].drop_duplicates().sort_values().map(lambda x: str(int(x)) if isinstance(x, (float, int)) else x)),
            'DISCOUNTR_MOM': (x.DISCOUNTR * x.ROLL_MIX).sum(),
            'DISCOUNTR': (x.DISCOUNTR * x.ROLL_MIX).sum(),
            'NATIONAL_SUBSIDY': '/'.join(x['NATIONAL_SUBSIDY'].drop_duplicates().sort_values().map(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, (float, int)) else '')),
            'WEIGHT_CP': (x.CP * x.ROLL_MIX).sum(),
        }))
    curr_data.DISCOUNTR_MOM = curr_data.DISCOUNTR_MOM - last_data.DISCOUNTR_MOM
    curr_data = curr_data[pd.notnull(curr_data.DISCOUNTR_MOM)]
    curr_data.insert(
        column='DISCOUNTR_MOM_DIFF',
        loc=list(curr_data).index('DISCOUNTR_MOM')+1,
        value=curr_data.DISCOUNTR_MOM * curr_data.WEIGHT_CP / 1000
    )
    curr_data.insert(
        column='DISCOUNTR_DIFF',
        loc=list(curr_data).index('DISCOUNTR')+1,
        value=curr_data.DISCOUNTR * curr_data.WEIGHT_CP / 1000
    )
    curr_data = curr_data.sort_values(['DISCOUNTR_MOM'])
    curr_data[['DISCOUNTR_MOM', 'DISCOUNTR']] = curr_data[['DISCOUNTR_MOM', 'DISCOUNTR']].applymap(lambda x: '{:.1%}'.format(normal_round(x, 3)))
    curr_data[['DISCOUNTR_MOM_DIFF', 'DISCOUNTR_DIFF']] = curr_data[['DISCOUNTR_MOM_DIFF', 'DISCOUNTR_DIFF']].applymap(lambda x: '{:.1f}'.format(normal_round(x, 1)).replace('.', ','))

    return curr_data.replace('-0,0', '0,0')


def AdjustSheetFormat(ws):

    ws.range('L:L').number_format = '0.0%'
    ws.range('N:N').number_format = '0.0%'
    ws.range(1, 1).column_width = 25

    data = ws.range('B1:G1').options(pd.DataFrame, index=False, expand='down').value
    filter_index = data[data.apply(lambda x: x.MSRP_MIN == x.CP_MIN and x.MSRP_MAX == x.CP_MAX, axis=1)].index.values
    for idx in filter_index:
        ws.range(f'B{idx+2}:G{idx+2}').color = colorRed
        ws.range(f'P{idx+2}').color = colorRed
    return None


def RGB(red, green, blue): 
    assert 0 <= red <= 255
    assert 0 <= green <= 255
    assert 0 <= blue <= 255
    return red + (green << 8) + (blue << 16)

colorNone, colorRed = RGB(255, 255, 255), RGB(255, 0, 0)


def normal_round(n, decimals=0):
    expoN = n * 10 ** decimals
    if abs(expoN) - abs(math.floor(expoN)) < 0.5:
        return math.floor(expoN) / 10 ** decimals
    return math.ceil(expoN) / 10 ** decimals
