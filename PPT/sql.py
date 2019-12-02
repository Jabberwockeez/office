import settings
import common
import pandas as pd
import xlwings as xw
import os


segment_list = ['小型轿车','紧凑型轿车','中型轿车','小型SUV','紧凑型SUV','中型SUV','小型MPV','中型MPV','大型MPV']      


def GetPriceIndexData(dimension):
    if 'TOTAL' in dimension:
        tmp_data = source_price_index_data
    else:
        tmp_data = source_price_index_data[source_price_index_data.CUST_SEGMENT_NAME_CHN.isin(segment_list)]

    tmp_data = tmp_data.assign(TOTAL='整体')\
                       .groupby(dimension)\
                       .apply(lambda x: pd.Series({'MSRP': (x.MSRP * x.SALES_QTY / x.SALES_QTY.sum()).sum(),
                                                   'TP': (x.AVG_TP * x.SALES_QTY / x.SALES_QTY.sum()).sum()}))
    tmp_data['DISCOUNT_RATE'] = tmp_data.apply(lambda x: (x.MSRP - x.TP) / x.MSRP, axis=1)
    return tmp_data.drop('MSRP', axis=1)


db = common.OracleDB('DBBG')
query = settings.ReadFile(settings.GetRecentFile(settings.sqlPath))\
    .replace('SPECIFICPRICE_DATE', str(settings.price_date))
source_price_index_data = db.query_data(query).query("PRICE_DATE == {}".format(settings.price_date))
price_index_data = GetPriceIndexData(['TOTAL']).append(GetPriceIndexData(['CUST_SEGMENT_NAME_CHN']))
print(price_index_data.shape)

price_level_data = db.query_data_from_procedure('pkg_report_changan.proc_price_level', [settings.price_date]).drop('PRICE_DATE', axis=1).set_index(['CUST_SEGMENT_NAME_CHN', 'PEICE_SEGMENT']).unstack('PEICE_SEGMENT').droplevel(level=0, axis=1)
print(price_level_data.shape)

manf_price_data = db.query_data_from_procedure('pkg_report_changan.proc_manf_wt_tp', [settings.price_date]).drop('PRICE_DATE', axis=1).set_index('MANF_NAME').rename(columns={'WEIGHT_TP':'TP', 'WEIGHT_DISCOUNT':'DISCOUNT_RATE'})
print(manf_price_data.shape)

db.close()
