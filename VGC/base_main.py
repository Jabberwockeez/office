import common
import settings
from base_process import CalculateVersionData, ReadVersionData, CalculateDimensionData, CalculateResultData, Write
import pandas as pd
import xlwings as xw
import os
import warnings
from sys import exit
warnings.filterwarnings('ignore')


dimension = ['BRAND', 'MODELGROUP', '报告车型名称', 'SEGMENT', 'BODYTYPE', 'PROPERTYTYPE', 'PROPERTY', 'BRANDNATION']
dt = 'Weekly_Report_' + pd.datetime.now().strftime('%Y%m%d%H%M')
result_dict = {
    f'子车型_{dt}.xlsx': 
        [['BRAND', '报告车型名称', 'SEGMENT', 'BRANDNATION'],
        ['BRAND', '报告车型名称', 'SEGMENT', 'PROPERTY', 'BRANDNATION']],
    f'品牌_{dt}.xlsx': 
        [['BRAND'],
        ['BRAND', 'PROPERTY']],
    f'Total_Pre_{dt}.xlsx': 
        [[],
        ['PROPERTYTYPE']],
    f'Total_Market_{dt}.xlsx': 
        [[],
        ['PROPERTY']],
}


if __name__ == "__main__":
    
    app = xw.App(add_book=False, visible=settings.isVisible)

    # Version
    ways_data = CalculateVersionData(dimension)
    columns_list = ('MSRP', 'TP', 'DISCOUNTR', 'ROLL_MIX')
    Write(app, ways_data, columns_list, f'VGC_Version_{dt}.xlsx')

    # Dimension
    vgc_data = CalculateDimensionData(ways_data, dimension)
    columns_list = ('MSRP', 'TP', 'DISCOUNTR', 'ROLL_DIMENSION_SALES', 'MSRP_INDEX')
    Write(app, vgc_data.query(f'PRICE_DATE >= {pd.datetime(settings.price_date//10000, 1, 1):%Y%m%d}'), columns_list, f'VGC_Dimension_{dt}.xlsx')

    # report
    for result_file, result_dimension_list in result_dict.items():
        if result_file.startswith('Total_Pre'):
            premium_brand = ['Acura','Audi','BMW','Infiniti','Jaguar','Land Rover','MB','Volvo','Lexus','Maserati','Mini','Porsche','Tesla']
            mask = vgc_data['BRAND'].isin(premium_brand)
            result_data = CalculateResultData(vgc_data[mask], result_dimension_list)
        else:
            result_data = CalculateResultData(vgc_data, result_dimension_list)
        columns_list = ('MSRP', 'TP', 'DISCOUNTR', 'MSRP_INDEX', 'TP_PI', 'DiscountR_环比')
        Write(app, result_data.query(f'PRICE_DATE >= {pd.datetime(settings.price_date//10000, 1, 1):%Y%m%d}'), columns_list, result_file)

    app.kill()
