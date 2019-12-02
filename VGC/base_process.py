import settings
import pandas as p8
from numpy import inf
import os
import xlwings as xw


def ShowDimension(dimension):
    if 'FUELTYPE' in dimension:
        return ['报告车型名称', 'BODYTYPE', 'PROPERTYTYPE', 'PROPERTY', 'FUELTYPE']
    else:
        return ['报告车型名称', 'BODYTYPE', 'PROPERTYTYPE', 'PROPERTY']


# 计算Version维度
def CalculateVersionData(dimension):
    from sql import rule_data, terminal_data, cpca_data, mix_data, tp_data, isabs_data

    keys = [*rule_data.columns, 'YM_ID']
    ways_data = \
        pd.concat([rule_data.query('MANF_PROP_NAME=="进口"').merge(terminal_data, on='SUB_MODEL_ID', how='left'), 
                   rule_data.query('MANF_PROP_NAME!="进口"').merge(cpca_data, on='SUB_MODEL_ID', how='left')]) \
          .merge(rule_data.merge(mix_data, on='VERSION_ID', how='left'), on=keys, how='outer') \
          .merge(rule_data.merge(tp_data, on='VERSION_ID', how='left'), on=keys, how='outer') \
          .merge(rule_data.merge(isabs_data, on='VERSION_ID', how='left'), on=keys, how='outer') \
          .dropna(subset=['YM_ID'])\
          .reset_index(drop=True)

    ways_data['MIX'] = \
        ways_data.groupby(['PRICE_DATE', 'SUB_MODEL_ID'])['MIX'].apply(lambda x: x / x.sum())
    roll_data = \
        ways_data[['VERSION_ID', 'YM_ID', 'VERSION_SALES']].drop_duplicates(subset=['VERSION_ID', 'YM_ID'])
    roll_data['ROLL_SALES'] = \
        roll_data.sort_values('YM_ID').groupby('VERSION_ID')['VERSION_SALES'].apply(lambda x: x.rolling(window=3, min_periods=1).sum().shift(1))    # 错月滚动
    ways_data = \
        ways_data.merge(roll_data.drop('VERSION_SALES', axis=1), on=list(roll_data)[:-2], how='left')
    ways_data['ROLL_MIX'] = \
        ways_data.query('TP > 0').groupby([*dimension, 'PRICE_DATE'])['ROLL_SALES'].apply(lambda x: x / x.sum())

    ways_data.insert(
        column='DISCOUNTR', 
        loc=list(ways_data).index('TP') + 1, 
        value=ways_data.query('TP > 0').apply(lambda x: x['TP'] / x['MSRP'] - 1 if x['MSRP'] > 0 else None, axis=1))
    ways_data.insert(
        column='SUB_MODEL_MIX_SALES', 
        loc=list(ways_data).index('MIX') + 1, 
        value=ways_data.apply(lambda x: x['SUB_MODEL_SALES'] * x['MIX'] if x['SUB_MODEL_SALES'] and x['MIX'] else None, axis=1))
    ways_data.insert(
        column='DIMENSION',
        loc=list(ways_data).index('VERSION_ID'),
        value=ways_data[ShowDimension(dimension)].apply(lambda x: ' '.join(x), axis=1))

    # 读取固定底层表
    last_data = ReadVersionData(dimension)
    if settings.isSQL:
        ways_data = ShiftVersionSales(ways_data, last_data, dimension)
    else:
        ways_data = ways_data.query('PRICE_DATE < 20190101').append(last_data)

    return ways_data.sort_values(by=['PRICE_DATE', 'BRAND', '报告车型名称', 'VERSION_ID'])


def ShiftVersionSales(ways_data, last_data, dimension):
    if not settings.isLastWeek:     # 不是最后一周,本周ROLL_SALES等于上周ROLL_SALES
        mask_last = last_data.PRICE_DATE == settings.last_price_date
        mask_ways = ways_data.PRICE_DATE == settings.price_date
        ways_data.loc[mask_ways, ['ROLL_SALES']] = ways_data.loc[mask_ways].reset_index().merge(last_data.loc[mask_last], on='VERSION_ID', how='left', suffixes=('_ways', '_last')).set_index('index')[['ROLL_SALES_last']].values
        ways_data.loc[mask_ways, 'ROLL_MIX'] = ways_data.loc[mask_ways].query('TP > 0').groupby([*dimension, 'PRICE_DATE'])['ROLL_SALES'].apply(lambda x: x / x.sum())
        return ways_data.query('PRICE_DATE < 20190101').append(last_data).append(ways_data.query('PRICE_DATE == @settings.price_date'))
    else:                           # 是最后一周,切换本月ROLL_SALES
        # 更新固定底层表内上个月的 SUB_MODEL_SALES, MIX, SUB_MODEL_MIX_SALES, VERSION_SALES
        mask_last = last_data.YM_ID == int(settings.last_month.strftime('%Y%m'))
        mask_ways = ways_data.YM_ID == int(settings.last_month.strftime('%Y%m'))
        last_data.loc[mask_last, ['SUB_MODEL_SALES', 'MIX', 'SUB_MODEL_MIX_SALES', 'VERSION_SALES']] = last_data.loc[mask_last].reset_index().merge(ways_data.loc[mask_ways], on=['VERSION_ID', 'PRICE_DATE'], how='left', suffixes=('_last', '_ways')).set_index('index')[['SUB_MODEL_SALES_ways', 'MIX_ways', 'SUB_MODEL_MIX_SALES_ways', 'VERSION_SALES_ways']].values
        # 更新固定底层表内本月的 ROLL_SALES, ROLL_MIX
        mask_last = last_data.YM_ID == int(settings.report_date.strftime('%Y%m'))
        mask_ways = ways_data.YM_ID == int(settings.report_date.strftime('%Y%m'))
        last_data.loc[mask_last, ['ROLL_SALES', 'ROLL_MIX']] = last_data.loc[mask_last].reset_index().merge(ways_data.loc[mask_ways], on=['VERSION_ID', 'PRICE_DATE'], suffixes=('_last', '_ways')).set_index('index')[['ROLL_SALES_ways', 'ROLL_MIX_ways']].values
        return ways_data.query('PRICE_DATE < 20190101').append(last_data).append(ways_data.query('PRICE_DATE == @settings.price_date'))


# 读取Version维度历史数
def ReadVersionData(dimension):
    from sql import rule_data

    file = settings.GetRecentFile(os.path.join(settings.sourcePath, 'Version_Base'), 'VGC_Version_Weekly')
    ways_data = pd.read_excel(file).loc[:, :'ROLL_MIX'].reset_index(drop=True)
    ways_data[['HALT_PRODUCT_DATE', 'HALT_SALE_DATE']] = ways_data.merge(rule_data[['VERSION_ID', 'HALT_PRODUCT_DATE', 'HALT_SALE_DATE']], on='VERSION_ID', how='left', suffixes=(' ', ''))[['HALT_PRODUCT_DATE', 'HALT_SALE_DATE']].fillna(ways_data)
    if 'FUELTYPE' in dimension:
        ways_data['ROLL_MIX'] = \
            ways_data.query('TP > 0').groupby([*dimension, 'PRICE_DATE'])['ROLL_MIX'].apply(lambda x: x / x.sum())

    return ways_data.sort_values(by=['PRICE_DATE', 'BRAND', '报告车型名称', 'VERSION_ID'])


# 计算Dimension维度
def CalculateDimensionData(ways_data, dimension):
    vgc_data = \
        ways_data.groupby([*dimension, 'YM_ID', 'PRICE_DATE']) \
                 .apply(lambda x: pd.Series({
                    'MSRP': (x['MSRP'] * x['ROLL_MIX']).sum()
                        if any(pd.notnull(x['MSRP'])) and x['ROLL_MIX'].sum() else None,
                    'TP': (x['TP'] * x['ROLL_MIX']).sum()
                        if any(pd.notnull(x['TP'])) and x['ROLL_MIX'].sum() else None,
                    'DISCOUNTR': (x['DISCOUNTR'] * x['ROLL_MIX']).sum()
                        if any(pd.notnull(x['DISCOUNTR'])) and x['ROLL_MIX'].sum() else None,
                    'DIMENSION_SALES': x['SUB_MODEL_MIX_SALES'].sum()
                        if x['SUB_MODEL_MIX_SALES'].sum() else None})) \
                 .reset_index()
    vgc_data['MSRP_INDEX'] = \
        vgc_data.sort_values('PRICE_DATE') \
                .groupby(dimension)['MSRP'] \
                .apply(lambda x: x.div(x.shift(1))) \
                .replace([inf,])
    vgc_data.insert(
        column='DIMENSION',
        loc=list(vgc_data).index('YM_ID'),
        value=vgc_data[ShowDimension(dimension)].apply(lambda x: ' '.join(x), axis=1))
    roll_data = \
        vgc_data[['DIMENSION', 'YM_ID', 'DIMENSION_SALES']].drop_duplicates(subset=['DIMENSION', 'YM_ID'])
    roll_data['ROLL_DIMENSION_SALES'] = \
        roll_data.sort_values('YM_ID').groupby(['DIMENSION'])['DIMENSION_SALES'].apply(lambda x: x.rolling(window=6, min_periods=1).sum().shift(1))
    vgc_data = \
        vgc_data.merge(roll_data.drop('DIMENSION_SALES', axis=1), on=list(roll_data)[:-2], how='left')

    modify_file = settings.GetRecentFile(os.path.join(settings.sourcePath, 'Modify'), '周报')
    sales_modify = \
        pd.read_excel(modify_file, r'ROLL_DIMENSION_SALES').filter(['DIMENSION', 'PRICE_DATE', 'ROLL_DIMENSION_SALES'], axis=1).drop_duplicates(['DIMENSION', 'PRICE_DATE'], keep='last')
    vgc_data = vgc_data.reset_index(drop=True)
    vgc_data = vgc_data.iloc[:, :-1].merge(sales_modify, on=list(sales_modify)[:-1], how='left').fillna(vgc_data)
    
    vgc_data = ShiftDimensionSales(vgc_data, dimension)

    return vgc_data.sort_values(by=['PRICE_DATE', 'BRAND', '报告车型名称'])


def ShiftDimensionSales(vgc_data, dimension):
    if not settings.isLastWeek:     # 不是最后一周,本月ROLL_SALES等于上周ROLL_SALES
        mask_last = vgc_data.PRICE_DATE == settings.last_month_price_date
        mask_curr = vgc_data.YM_ID == settings.price_date // 100
        vgc_data.loc[mask_curr, ['ROLL_DIMENSION_SALES']] = \
            vgc_data.loc[mask_curr].reset_index().merge(vgc_data.loc[mask_last], on=dimension, how='left', suffixes=('_curr', '_last')).set_index('index')[['ROLL_DIMENSION_SALES_last']].values
        return vgc_data
    else:                           # 是最后一周,切换本月ROLL_SALES
        return vgc_data


# 计算结果
def CalculateResultData(vgc_data, result_dimension_list):
    result_data = pd.DataFrame()
    for result_dimension in result_dimension_list:
        vgc_data['ROLL_DIMENSION_MIX'] = \
            vgc_data.query('TP > 0').groupby([*result_dimension, 'PRICE_DATE'])['ROLL_DIMENSION_SALES'].apply(lambda x: x / x.sum())
        tmp_data = \
            vgc_data.query('ROLL_DIMENSION_MIX > 0')\
                    .groupby([*result_dimension, 'YM_ID', 'PRICE_DATE']) \
                    .apply(lambda x: pd.Series({
                        'MSRP': (x['MSRP'] * x['ROLL_DIMENSION_MIX']).sum()
                            if any(pd.notnull(x['MSRP'])) and x['ROLL_DIMENSION_MIX'].sum() else None,
                        'TP': (x['TP'] * x['ROLL_DIMENSION_MIX']).sum()
                            if any(pd.notnull(x['TP'])) and x['ROLL_DIMENSION_MIX'].sum() else None,
                        'DISCOUNTR': (x['DISCOUNTR'] * x['ROLL_DIMENSION_MIX']).sum()
                            if any(pd.notnull(x['DISCOUNTR'])) and x['ROLL_DIMENSION_MIX'].sum() else None,
                        'MSRP_INDEX': (x['MSRP_INDEX'] * x['ROLL_DIMENSION_MIX']).sum()
                            if any(pd.notnull(x['MSRP_INDEX'])) and x['ROLL_DIMENSION_MIX'].sum() else None,
                        'DIMENSION_SALES': x['ROLL_DIMENSION_SALES'].sum()
                            if x['ROLL_DIMENSION_SALES'].sum() else None})) \
                    .reset_index()
        tmp_data['TP_PI'] = \
            tmp_data.apply(lambda x: x['DISCOUNTR'] + 1 if pd.notnull(x['DISCOUNTR']) else nan, axis=1)

        if len(result_dimension):
            tmp_data['DiscountR_环比'] = tmp_data.groupby(result_dimension)['DISCOUNTR'].apply(lambda x: x.diff(periods=1))
        else:
            tmp_data['DiscountR_环比'] = tmp_data[['DISCOUNTR']].apply(lambda x: x.diff(periods=1))

        if len(result_data):
            placeholder_col = list(set(tmp_data.columns) - set(result_data.columns))[0]
            result_data[placeholder_col] = 'Total'
        result_data = result_data.append(tmp_data).reindex(tmp_data.columns, axis=1)

    return result_data


# 生成Excel
def Write(app, data, columns_list, filename):
    data.to_excel(os.path.join(settings.reportBasePath, '竖版_'+filename), index=False)
    wb = app.books.add()
    index_col = [*list(data)[:list(data).index('YM_ID')+1], 'PRICE_DATE']
    data = data.assign(PRICE_DATE=data.PRICE_DATE.map(fmt_week)).set_index(index_col)
    for col in columns_list:
        ws = wb.sheets.add(col, after=wb.sheets.count)
        ws.range(1, 1).value = data.loc[:, col:col].unstack(['YM_ID', 'PRICE_DATE'])
        ws.range(4, len(index_col)-1).select()
        app.api.Windows(wb.name).FreezePanes = True
    wb.sheets[0].delete()
    wb.save(os.path.join(settings.reportBasePath, '横版_'+filename))
    wb.close()
    return None


def fmt_week(x):
    if x % 100 == 8:
        return 'W1'
    elif x % 100 == 15:
        return 'W2'
    elif x % 100 == 22:
        return 'W3'
    elif x % 100 == 28 or x % 100 == 30:
        return 'W4'
    elif x % 100 == 31 or x % 100 == 3:
        return 'W5'
    else:
        return 'Error'