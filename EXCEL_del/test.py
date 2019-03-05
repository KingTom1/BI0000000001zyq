import EXCEL_del.read_excle as read_excle
import pandas as pd
# 自定义路径
def Get_dict():
    excel_path = r"excels\规则.xlsx"
    excel_path1 = r"excels\病人来源201810-12.xlsx"
    sheets =['CDS-成都市','SCS-四川省','外省']
    sheet1 =u'门诊201810-12'
    j = 3
    # 获取类对象
    arr = []
    for sheet in sheets:
        RE = read_excle.ReadTwoExcel(excel_path, sheet)
        a = RE.read_cow(excel_path,sheet,0)
        # print(a)
        b = RE.read_cow(excel_path1,sheet1,j)
        # print(b)
        c = dict()
        for a_ in a:
            for b_ in b:
                # print(b_.find(a_))
                if b_.find(a_)>0:
                    c[a_] = b_
        arr.append(c)
        j = j-1
    return arr
# a = Get_dict()
# print(a[1])
# df = RE.readExcel(excel_path1,sheet1)
# # print(df)
# pv1 = pd.pivot_table(df,aggfunc='sum',values='就诊人次',index=a[0])
# # print(pv1.loc['ABX-阿坝县'])
# d = dict()
# for k in c.keys():
#     d[k]=pv1.ix[c[k]]
# # print(d)
# s = pd.DataFrame.from_dict(d,orient='index')
# s = s.sort_values('就诊人次',ascending=False)
# total = sum(s['就诊人次'])
# s.loc['合计'] = s.apply(lambda x: x.sum())
# print(total)
# s_ = s.reset_index()
# print(s_)

# print('办卡地址区县'.find('办卡地址区县'))   输出 0



# df = RE.readExcel(r"excels\2018年季度门诊.xlsx",u'统计结果')
# print(df)
# import xlrd
# excel_path3=r'excels\统计结果表格式参照.xlsx'
# sheet3 = r'Sheet1'
# book = xlrd.open_workbook(excel_path3)#打开一个excel
# # sheet = book.sheet_by_index(0)#根据顺序获取sheet
# sheet = book.sheet_by_name(sheet3)#根据sheet页名字获取sheet
# print(sheet.cell(0,0).value)#指定行和列获取数据



# 获取与第二张表的对应关系
# zidian = RE.zidian("门诊诊疗费用明细记录关联门诊费用")
# print(zidian)

# 实验室检验记录明细记录关联实验室记录
# 门诊诊疗费用明细记录关联门诊费用

#
# sql = "SELECT T1.ORGNAME 医院名称,       T1.ORGCODE 机构代码,       T1.GLSL 关联数据量,       T2.BSJL 表数据量,       ROUND((convert(float,T1.GLSL) / convert(float,T2.BSJL)) * 100, 2)  关联度  FROM (SELECT T.ORGANIZATION_NAME ORGNAME,               T.ORGANIZATION_CODE ORGCODE,               COUNT(*) GLSL          FROM DI_ADI_REGISTER_INFO T         WHERE substring(T.DATAGENERATE_DATE, 1, 4) = '2018'                     AND (T.ORGANIZATION_CODE+ T.BUSINESS_ID+ T.LOCAL_ID) IN               (SELECT ORGANIZATION_CODE+ BUSINESS_ID+ LOCAL_ID                  FROM DI_PATIENT_TREAT_INFO)         GROUP BY T.ORGANIZATION_NAME, T.ORGANIZATION_CODE) T1  LEFT JOIN (SELECT T.ORGANIZATION_NAME ORGNAME,                    T.ORGANIZATION_CODE ORGCODE,                    COUNT(*) BSJL               FROM DI_ADI_REGISTER_INFO T              WHERE substring(T.DATAGENERATE_DATE, 1, 4) = '2018'                             GROUP BY T.ORGANIZATION_NAME, T.ORGANIZATION_CODE) T2    ON T1.ORGCODE = T2.ORGCODE"
# print(sql.replace("2017","2016"))