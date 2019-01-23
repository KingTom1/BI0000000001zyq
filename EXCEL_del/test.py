import EXCEL_del.read_excle as read_excle

# 自定义路径
excel_path = r"excels\规则.xlsx"
excel_path1 = r"excels\病人来源201810-12.xlsx"
sheet =u'规则'
sheet1 =u'门诊201810-12'
col_name1 = '指标项'
col_sql = '统计SQL'
# 获取类对象
RE = read_excle.ReadTwoExcel(excel_path, excel_path1, sheet, sheet1,col_name1,col_sql)
a = RE.read_cow()
print(a)

# 获取与第二张表的对应关系
# zidian = RE.zidian("门诊诊疗费用明细记录关联门诊费用")
# print(zidian)

# 实验室检验记录明细记录关联实验室记录
# 门诊诊疗费用明细记录关联门诊费用

#
# sql = "SELECT T1.ORGNAME 医院名称,       T1.ORGCODE 机构代码,       T1.GLSL 关联数据量,       T2.BSJL 表数据量,       ROUND((convert(float,T1.GLSL) / convert(float,T2.BSJL)) * 100, 2)  关联度  FROM (SELECT T.ORGANIZATION_NAME ORGNAME,               T.ORGANIZATION_CODE ORGCODE,               COUNT(*) GLSL          FROM DI_ADI_REGISTER_INFO T         WHERE substring(T.DATAGENERATE_DATE, 1, 4) = '2018'                     AND (T.ORGANIZATION_CODE+ T.BUSINESS_ID+ T.LOCAL_ID) IN               (SELECT ORGANIZATION_CODE+ BUSINESS_ID+ LOCAL_ID                  FROM DI_PATIENT_TREAT_INFO)         GROUP BY T.ORGANIZATION_NAME, T.ORGANIZATION_CODE) T1  LEFT JOIN (SELECT T.ORGANIZATION_NAME ORGNAME,                    T.ORGANIZATION_CODE ORGCODE,                    COUNT(*) BSJL               FROM DI_ADI_REGISTER_INFO T              WHERE substring(T.DATAGENERATE_DATE, 1, 4) = '2018'                             GROUP BY T.ORGANIZATION_NAME, T.ORGANIZATION_CODE) T2    ON T1.ORGCODE = T2.ORGCODE"
# print(sql.replace("2017","2016"))