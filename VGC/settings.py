'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-02-11 18:37:15
'''

import common
import os
import datetime
from dateutil.relativedelta import relativedelta
import logging


isVisible = 0
isLastWeek = 0
isSQL = 0

# 时间
report_date = datetime.datetime.now()
# report_date = datetime.datetime(2019, 10, 30)
last_month = report_date + relativedelta(months=-1)

price_date = int(common.GetReportDate(report_date, '周报').get('DATE_NAME'))
print('TP日期: ', price_date)
last_price_date = int(common.GetReportDate(price_date - 1, '周报').get('DATE_NAME'))
print('上周TP日期: ', last_price_date)
last_month_price_date = int(common.GetReportDate(report_date.replace(day=1), '周报').get('DATE_NAME'))
print('上个月最后一周TP日期: ', last_month_price_date)

if isLastWeek:
    mix_date = int('{:%Y%m}'.format(last_month))
else:
    mix_date = int('{:%Y%m}'.format(last_month + relativedelta(months=-1)))
print('MIX日期: ', mix_date)


# 路径
path = os.path.dirname(os.path.dirname(__file__))
sourcePath = os.path.join(path, r'Source')
reportBasePath = os.path.join(path, r'Reports_Base')


# 获取指定目录下最新文件
def GetRecentFile(path, key=''):
    file_filter = filter(lambda x: key in x, os.listdir(path))
    file = sorted(file_filter, key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)[0]
    print(os.path.join(path, file))
    return os.path.join(path, file)


# # 日志设置
# logger = logging.getLogger(__name__)
# logger.setLevel(logging.INFO)

# logfile = os.path.join(path, 'record.log')
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y/%m/%d/ %H:%M:%S')
# handler = logging.FileHandler(logfile)
# handler.setLevel(logging.INFO)
# handler.setFormatter(formatter)
# logger.addHandler(handler)
