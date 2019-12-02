import settings
import common
import os
import xlwings as xw
import pandas as pd


db = common.OracleDB('DBBG')

brand_data = db.query_data(
    f"""
    SELECT T.VERSION_ID, T.BRAND_ENAME
      FROM SJZX.V_VERSION T
    """
)

db.close()