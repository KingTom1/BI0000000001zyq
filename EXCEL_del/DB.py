import pyodbc
import pandas as pd
# 连接数据库,获取连接流对象
def connect(cf):
    SERVER = cf.get('db', 'SERVER')
    DATABASE = cf.get('db', 'DATABASE')
    UID = cf.get('db', 'UID')
    PWD = cf.get('db','PWD')
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + UID + ';PWD='+ PWD +'')
    cursor = cnxn.cursor()
    return cursor

# 查询sql
def get_result(cursor,sql):
    cursor.execute("SELECT @@version;")
    cursor.execute(sql)
    rows = cursor.fetchall()
    Hospit = [row[0] for row in rows]
    Hoscode = [row[1] for row in rows]
    KeyData = [row[2] for row in rows]
    TableData = [row[3] for row in rows]
    KeyRate = [row[4] for row in rows]
    df= pd.DataFrame({
        '医院名称': Hospit,
        '机构代码': Hoscode,
        '关联数据量': KeyData,
        '表数据量': TableData,
        '关联度': KeyRate,
        })
    return df
