'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-02-11 18:37:15
'''

import common
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import logging


isVisible = 0

# 时间
report_date = datetime.now()
# report_date = datetime(2019, 10, 29)
last_month = report_date + relativedelta(months=-1, days=-1)

# 价格日期
price_date = int(common.GetReportDate(report_date, '月报').get('DATE_NAME'))
mix_date = int(common.GetReportDate(last_month, '月报').get('DATE_NAME') // 100)
print(f'TP日期: {price_date}\nMIX日期: {mix_date}')

# 路径
path = os.path.dirname(os.path.dirname(__file__))
sourcePath = os.path.join(path, r'Source')
reportPath = os.path.join(path, r'Reports')
sqlPath = os.path.join(path, r'SQL')


def GetRecentFile(path, key=''):
    file_filter = filter(lambda x: key in x and '~' not in x, os.listdir(path))
    file = sorted(file_filter, key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)[0]
    print(file)
    return os.path.join(path, file)


# 读取文本文件
def ReadFile(filepath):
    query = ''
    with open(filepath, 'r', encoding='utf8') as f:
        for line in f:
            query += line
    return query


# 日志设置
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

logfile = os.path.join(path, f'record_{datetime.now():%Y%m%d%H%M}.log')
formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%Y/%m/%d/ %H:%M:%S')
handler = logging.FileHandler(logfile)
handler.setLevel(logging.INFO)
handler.setFormatter(formatter)
logger.addHandler(handler)
