# coding: utf-8
'''
@Author: Guo Yulong
@Date: 2019-01-06 18:28:20
@LastEditors: Guo Yulong
@LastEditTime: 2019-02-11 18:37:15
'''

import os
# import logging


isVisible = 1

jsonFileName = r'调整.json'
sourceFileName = r'测试问卷报告样板.docx'


# 路径
path = os.path.dirname(os.path.dirname(__file__))
sourcePath = os.path.join(path, r'Source')
reportPath = os.path.join(path, r'Reports')
sqlPath = os.path.join(path, r'SQL')

# # 日志设置
# logger = logging.getLogger(__name__)
# logger.setLevel(logging.INFO)

# logfile = os.path.join(path, 'record.log')
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y/%m/%d/ %H:%M:%S')
# handler = logging.FileHandler(logfile)
# handler.setLevel(logging.INFO)
# handler.setFormatter(formatter)
# logger.addHandler(handler)
