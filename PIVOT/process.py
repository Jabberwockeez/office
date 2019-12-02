import settings
import xlwings as xw
import pandas as pd
import os



def MergeData(data):

    columns = list(data)
    groupby = [col for col in columns if col not in ('厂商', '品牌')]
    tmp_data = data.groupby(groupby, as_index=False)\
        .apply(lambda x: pd.Series({'厂商': '、'.join(JoinSplit(x.厂商)),'品牌': '、'.join(JoinSplit(x.品牌))}))\
        .reset_index()\
        .reindex(columns=data.columns)
    return tmp_data


def PivotData(data):

    columns = list(data)
    groupby = [col for col in columns if col not in ('厂商', '品牌')]
    tmp_data = data.groupby('经销商集团').apply(lambda x: pd.Series({
        '广丰店数': '、'.join(x.厂商).split('、').count('广汽丰田'),
        '分布省份数量': x.省份.drop_duplicates().count(),
        '品牌数量': len(JoinSplit(x.品牌)),
        '4S店数量': x.drop(['厂商', '品牌'], axis=1).drop_duplicates().shape[0],
        '品牌覆盖': '、'.join(JoinSplit(x.品牌)),
        '省份覆盖': CombineRegion(x),
    })).reset_index()
    tmp_data['集团简介'] = tmp_data.apply(lambda x: f"集团覆盖{x.省份覆盖}等{x.分布省份数量}个省份，经营的品牌包含{x.品牌覆盖}等{x.品牌数量}个品牌，全国共{x['4S店数量']}个4S店。", axis=1)
    return tmp_data


def MergeRankData(data):

    index_col = list(data)[:-3]
    pivot_col = list(data)[-3:]
    data['出资比例 '] = data['出资比例 '].map(lambda x: float(x.strip('%'))/100 if isinstance(x, str) else x)
    data['rank'] = data.groupby(index_col)['出资比例 '].rank(ascending=False, method='first')
    data = data.groupby(index_col)\
               .apply(lambda x: (x.sort_values(by=['rank'])).head(5))\
               .reset_index(drop=True)\
               .set_index([*index_col, 'rank'])\
               .unstack('rank')\
               .swaplevel(axis=1)\
               .sort_index(axis=1)\
               .reindex(pivot_col, axis=1, level=1)\
               .reset_index()
    return data


def _CombineRegion(df):
    """按省份统计城市信息.
    省份数量小于3个（含3个），需要显示省份旗下的城市；如省份数量大于3个，只需要显示省份即可。
    Args:
        df: pd.DataFrame.
    Returns:
        str. 举例: 
        省份数量大于3个 -> 广东省、天津市、浙江省、辽宁省、上海市、内蒙古自治区、河北省
        省份数量小于3个 -> 上海市、江苏省(南京市)、浙江省(绍兴市) 
    """
    if df.省份.drop_duplicates().count() > 3:
        return '、'.join(df.省份.drop_duplicates())
    else:
        tmp = df.groupby('省份').apply(lambda x: '、'.join(x.城市.drop_duplicates())).reset_index(name='城市')
        return '、'.join(tmp.apply(lambda x: f'{x.省份}({x.城市})' if not x.省份 == x.城市 else x.省份, axis=1))


def CombineRegion(df):
    """按省份统计城市信息.
    Args:
        df: pd.DataFrame.
    Returns:
        str. 福建省(1个城市）、贵州省(1个城市）、江苏省(2个城市）、江西省(4个城市）
    """
    tmp = df.groupby('省份').apply(lambda x: x.城市.drop_duplicates().count()).reset_index(name='城市')
    return '、'.join(tmp.apply(lambda x: f'{x.省份}({x.城市}个城市)', axis=1))


def JoinSplit(x):
    """先将序列连接成字符串, 再拆分, 最后去重.
    Args:
        x: pd.Series, DataFrame的一行或一列
    Returns:
        list. 按每个元素第一次出现的顺序. 举例: ['丰田', '现代、丰田', '宝马', '大众'] -> ['丰田', '现代', '宝马', '大众']
    """
    
    return unique('、'.join(x.drop_duplicates()).split('、'))


def unique(seq):
    """对序列去重并保持顺序不变.
    Args:
        seq: 任何可迭代序列, list, tuple, np.array, pd.Series...
    Returns:
        list. 按每个元素第一次出现的顺序.
    """

    seen = set()
    seen_add = seen.add
    return [x for x in seq if not (x in seen or seen_add(x))]