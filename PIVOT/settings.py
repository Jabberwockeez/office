'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-02-11 18:37:15
'''

import os
import datetime
from dateutil.relativedelta import relativedelta
import logging


isVisible = 0

# 时间
date_now = datetime.datetime.now() + relativedelta(months=-1)
# date_now = datetime.datetime(2019, 3, 30)

# 路径
path = os.path.dirname(os.path.dirname(__file__))
sourcePath = os.path.join(path, r'Source')
reportPath = os.path.join(path, r'Reports')



# 获取指定目录下最新文件
def GetRecentFile(path, key=''):
    file_filter = filter(lambda x: key in x and '~' not in x, os.listdir(path))
    file = sorted(file_filter, key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)[0]
    print(file)
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
