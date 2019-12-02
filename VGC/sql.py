import settings
import common
import pandas as pd
import os

db = common.OracleDB('DBBG')

rule_data = db.query_data(
    f"""
    SELECT T1.BRAND, T1.MODELNAME, T1.MODELGROUP, T1.报告车型名称, T1.SEGMENT, T1.BODYTYPE, T1.FUELTYPE,
           T1.PROPERTYTYPE, T1.PROPERTY, T1.BRANDNATION, T1.VERSION_ID,
           T2.VERSION_FULL_NAME, T2.LAUNCH_DATE, T2.HALT_PRODUCT_DATE, T2.HALT_SALE_DATE, T2.SUB_MODEL_ID, T2.MANF_PROP_NAME
      FROM T_RULE_VGC T1
      LEFT JOIN V_VERSION T2 ON T1.VERSION_ID = T2.VERSION_ID
    """
    )
print(rule_data.shape)

terminal_data = db.query_data(
    f"""
    SELECT T.YM_ID, T.SUB_MODEL_ID, SUM(T.SALES_QTY) SUB_MODEL_SALES
      FROM FDW_SUB_MODEL_SALES T
     WHERE 1 = 1 
       AND T.YM_ID >= {(settings.mix_date//100-1)*100+1} 
       AND T.YM_ID <= {settings.mix_date}
       AND T.DATA_SOURCE_ID = 5 
       AND T.SPACE_TYPE_ID = 5
     GROUP BY T.YM_ID, T.SUB_MODEL_ID
    """
    ).query("SUB_MODEL_ID in @rule_data.SUB_MODEL_ID")
print(terminal_data.shape)

cpca_data = db.query_data(
    f"""
    SELECT T.YM_ID, T.SUB_MODEL_ID, T.SALES_QTY SUB_MODEL_SALES
      FROM FDW_SUB_MODEL_SALES T
     WHERE 1 = 1 
       AND T.YM_ID >= {(settings.mix_date//100-1)*100+1} 
       AND T.YM_ID <= {settings.mix_date}
       AND T.DATA_SOURCE_ID = 3 
       AND T.SPACE_TYPE_ID = 2
    """
    ).query("SUB_MODEL_ID in @rule_data.SUB_MODEL_ID")
print(cpca_data.shape)

mix_data = db.query_data(
    f"""
    SELECT T.YM_ID, T.VERSION_ID, T.MIX
      FROM FDW_SUB_MODEL_MIX T
     WHERE 1 = 1
       AND T.YM_ID >= {(settings.mix_date//100-1)*100+1} 
       AND T.YM_ID <= {settings.mix_date}
       AND T.DATA_SOURCE_ID = 3 
       AND T.SPACE_TYPE_ID = 2
    """
    ).query("VERSION_ID in @rule_data.VERSION_ID")
print(mix_data.shape)

tp_data = db.query_data(
    f"""
    SELECT TRUNC(T.PRICE_DATE/100) YM_ID, T.PRICE_DATE, T.VERSION_ID, T.MSRP*10000 MSRP, AVG(T.LOWEST_PRICE)*10000 TP
      FROM V_TP_CITY T
      LEFT JOIN DM_CITY T1 ON T.CITY_ID = T1.CITY_ID
     WHERE 1 = 1
       AND T.PRICE_DATE >= {(settings.mix_date//100-1)*100+1}00 
       AND T.MSRP > 0
       AND T1.CITY_NAME IN ('北京市','上海市','广州市','深圳市','成都市','武汉市','沈阳市','杭州市','长沙市','西安市')
     GROUP BY T.VERSION_ID, T.PRICE_DATE, T.MSRP
    """
    ).query("VERSION_ID in @rule_data.VERSION_ID")
print(tp_data.shape)

isabs_data = db.query_data(
    f"""
    SELECT T.YM_ID, T.VERSION_ID, T.SALES_QTY VERSION_SALES
      FROM V_ISABS_VERSION_SALES T
     WHERE 1 = 1
       AND T.YM_ID >= {(settings.mix_date//100-1)*100+1}
       AND T.YM_ID <= {settings.mix_date}
    """
    ).query("VERSION_ID in @rule_data.VERSION_ID")
print(isabs_data.shape)

db.close()


# modify_file = settings.GetRecentFile(os.path.join(settings.sourcePath, 'Modify'), '周报')

# tp_modify = pd.read_excel(modify_file, r'TP')\
#               .filter(['VERSION_ID', 'PRICE_DATE', 'TP'], axis=1)\
#               .drop_duplicates(['VERSION_ID', 'PRICE_DATE'], keep='last')
# tp_data = tp_data.reset_index(drop=True)
# tp_data = tp_data.iloc[:, :-1].merge(tp_modify, on=list(tp_modify)[:-1], how='left').fillna(tp_data)
