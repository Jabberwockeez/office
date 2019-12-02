""" 
 * @Author: Guo Yulong 
 * @Date: 2018-08-24 20:11:12 
 * @Last Modified by: Guo Yulong
 * @Last Modified time: 2018-09-28 10:53:22
"""

import cx_Oracle
import pyodbc
from pandas import read_sql, DataFrame, Period, datetime


# Oracle 对象
class OracleDB(object):

    def __init__(self, DataBase):
        self.DataBase = DataBase
        
        if self.DataBase == 'DBBG':
            self.ip, self.port, self.SID = '172.16.1.52', 1521, self.DataBase
            self.username, self.password = 'SJZX', 'SJZX#2016'
        elif self.DataBase == 'DBDM':
            self.ip, self.port, self.SID = '172.16.1.99', 1521, self.DataBase
            self.username, self.password = 'SGMDM', 'bk9k74zi'
        elif self.DataBase == 'DBND':
            self.ip, self.port, self.SID = '172.16.1.55', 1521, self.DataBase
            self.username, self.password = 'SJZX', 'Sjzx#369'
        else:
            raise Exception('{} 不存在，无法建立连接。'.format(self.DataBase))

        self.dsn_tns = cx_Oracle.makedsn(self.ip, self.port, self.SID)
        self.conn = cx_Oracle.connect(self.username, self.password, self.dsn_tns)

    def query_data(self, query):
        self.data = read_sql(query, con = self.conn)
        return self.data

    def query_data_from_table(self, table):
        self.data = read_sql('SELECT * FROM {}'.format(table), con = self.conn)
        return self.data

    def query_data_from_procedure(self, procedure, param_list):
        self.cur = self.conn.cursor()
        self.crsr = self.cur.var(cx_Oracle.CURSOR)
        self.param = list(map(str, param_list)) + [self.crsr]
        self.cur.callproc(procedure, self.param)
        self.columns = [col[0] for col in self.crsr.getvalue().description]
        self.values = self.crsr.getvalue().fetchall()
        self.cur.close()
        return DataFrame(self.values, columns = self.columns)

    def execute_query(self, query):
        self.cur = self.conn.cursor()
        self.cur.execute(query)
        self.cur.close()
        return 0

    def close(self):
        self.conn.close()


# SQL 对象
class SQLDB(object):
    
    def __init__(self):
        self.driver, self.server, self.database, self.uid, self.pwd = '{SQL Native Client}', '172.16.1.2', 'test', 'dc_read', 'read!@#321'
        self.conn = pyodbc.connect(driver = self.driver, server = self.server, database = self.database, uid = self.uid, pwd = self.pwd)

    def query_data(self, query):
        self.data = read_sql(query, con = self.conn)
        return self.data

    def close(self):
        self.conn.close()


# 获取报告日期
""" 
date: int 当前8位日期
type: str 报告类型(周报,月报,半月报),默认周报
"""
def GetReportDate(date, type='周报'):

    if isinstance(date, (int, float)):
        pass
    elif isinstance(date, datetime):
        date = int(date.strftime('%Y%m%d'))
    elif isinstance(date, str):
        date = int(date)
    else:
        raise TypeError("date must be a int or string or datetime.")

    if type == '周报':
        db = OracleDB('DBBG')
        date_table = db.query_data_from_table('T_FA_DATE')    # DM_GATH_TIME
        db.close()
        mask = (date_table['DATE_TYPE'] == '报告周期') & (date_table['DATE_NAME'] <= date)
        return date_table[mask].iloc[-1].to_dict()
    elif type == '半月报':
        db = OracleDB('DBBG')
        date_table = db.query_data_from_table('T_FA_DATE')    # DM_GATH_TIME
        db.close()
        mask = (date_table['DATE_TYPE'] == '报告日期') & (date_table['DATE_NAME'] <= date)
        return date_table[mask].iloc[-1].to_dict()
    elif type == '月报':
        db = OracleDB('DBBG')
        date_table = db.query_data_from_table('T_FA_DATE')    # DM_GATH_TIME
        db.close()
        mask = (date_table['DATE_TYPE'] == '报告日期') & (date_table['DATE_VAL'].str.contains('下')) & (date_table['DATE_NAME'] <= date)
        return date_table[mask].iloc[-1].to_dict()
    else:
        pass


def YMIDToYearMonth(x):
    x = int(x)
    return Period(year = x // 100, month = x % 100, freq = 'M')


def GetLatestMixDate():
    db = OracleDB('DBBG')
    mix_date = db.query_data('SELECT MAX(YM_ID) FROM SJZX.FDW_SUB_MODEL_MIX')
    db.close()
    return int(mix_date.iloc[0, 0])
