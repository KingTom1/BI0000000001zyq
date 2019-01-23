import pandas as pd
import configparser
import EXCEL_del.DB as DB
import EXCEL_del.read_excle as read_excle
if __name__ == '__main__':
    # 自定义路径
    excel_path = r"excels\规则.xlsx"
    excel_path1 = r"excels\病人来源201810-12.xlsx"
    sheet = u'规则'
    sheet1 = u'门诊201810-12'
    col_name1 = '指标项'
    col_sql = '统计SQL'
    years = ['2016', '2017', '2018']
    # 获取类对象
    RE = read_excle.ReadTwoExcel(excel_path, excel_path1, sheet, sheet1, col_name1, col_sql)
    # 获取第一张表的第一行数据
    firstrow_data = RE.read_cow()
    d = RE.readExcel(excel_path, sheet)
    # 处理第一张表的字段
    KeyCodes = RE.GetKeyCode(firstrow_data, d)
    # 获取与第二张表的对应关系
    zidian = RE.zidian(KeyCodes)
    # 处理字典，取值组成数组
    arr = RE.arr(zidian)
    # 得到sql数组
    d1 = RE.readExcel(excel_path1, sheet1)
    # 读配置文件
    cf = configparser.ConfigParser()
    cf.read("cof.conf", 'utf-8')
    # 获取数据库对象
    cursor = DB.connect(cf)
    # 获取所有sqls语句
    sqls = RE.get_sql_strs(d1,arr)

    writer = pd.ExcelWriter("sqls.xlsx")
    sqls.to_excel(writer, sheet_name="sheet1")
    writer.save()

    # 获取第一个Excel关键字段
    arr_k = RE.arr_k(zidian)
    # 返回查询结果，并存储到新Excel
    for year in years:
        df_arr = RE.run_sql(sqls, zidian, year, cursor, arr_k, arr)
        writer = pd.ExcelWriter(year + ".xlsx")
        df_arr.to_excel(writer, sheet_name=year)
        writer.save()
