'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-01-28 13:25:50
'''

import os
import datetime
from dateutil.relativedelta import relativedelta
import win32com.client as com
com.gencache.EnsureDispatch('Excel.Application')


isVisible = 1

# 路径
path = os.path.dirname(os.path.dirname(__file__))
origPath = os.path.join(path, r'Orig')
sourcePath = os.path.join(path, r'Source')
resultPath = os.path.join(path, r'Reports')


def GetRecentFile(path, key=''):
    file_filter = filter(lambda x: key in x and '~' not in x, os.listdir(path))
    file = sorted(file_filter, key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)[0]
    print('打开文件: ', file)
    return os.path.join(path, file)
