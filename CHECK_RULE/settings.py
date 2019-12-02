'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-02-11 18:37:15
'''

import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import logging


isVisible = 0


# 路径
path = os.path.dirname(os.path.dirname(__file__))
finishPath = os.path.join(path, r'Finish')
sourcePath = os.path.join(path, r'Source')
reportPath = os.path.join(path, r'Reports')


# 日志设置
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

logfile = os.path.join(path, 'record.log')
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y/%m/%d/ %H:%M:%S')
handler = logging.FileHandler(logfile)
handler.setLevel(logging.INFO)
handler.setFormatter(formatter)
logger.addHandler(handler)
