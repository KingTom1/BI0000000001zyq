import pandas as pd
import EXCEL_del.test as t
import EXCEL_del.read_excle as read_excle
if __name__ == '__main__':
    # 自定义路径
    excel_path = r"excels\规则.xlsx"
    excel_path1 = r"excels\病人来源门诊统计2014-2018.xlsx"
    sheets = ['CDS-成都市','SCS-四川省','外省']
    sheet1 = u'Sheet1'
    COL = 1
    # 获取数据源数据
    df = pd.read_excel(excel_path1, sheet_name=sheet1)
    arr = t.Get_dict()
    writer = pd.ExcelWriter("%s统计结果.xlsx" % excel_path1)
    df_1 = df[df['办卡地址省']==sheets[1]]
    df_2 = df_1[df_1['办卡地址市']==sheets[0]]
    total_df2 = df_2.pivot_table(index='年',values='就诊人次',aggfunc='sum')
    # print(total_df2)
    result0= pd.DataFrame()
    for arr_ in arr[0].values():
        pv1 = df_2[df_2['办卡地址区县']==arr_]
        result0 = pv1.append(result0)

    total_result0 = result0.pivot_table(index='年', values='就诊人次', aggfunc='sum')
    # print(total_result0)
    new_df = total_df2-total_result0
    for index,row in new_df.iterrows():
        new = pd.DataFrame({"年": index, '办卡地址省': '四川省', '办卡地址市': "成都市", '办卡地址区县': '其他', '就诊人次': row})
        result0 = result0.append(new,ignore_index=True)
    print(result0)
    pv_result = pd.pivot_table(result0,index=['年','办卡地址省','办卡地址市','办卡地址区县'],values='就诊人次',aggfunc='sum')
    print(pv_result)
    pv_result.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)

    # COL = COL+6
    # result1 = pd.DataFrame()
    # for arr_ in arr[1].values():
    #     pv1 = df_1[df_1['办卡地址市']==arr_]
    #     result1 = pv1.append(result1)
    # pv_result1 = pd.pivot_table(result1, index=['年', '办卡地址省', '办卡地址市'], values='就诊人次',aggfunc='sum')
    # pv_result1.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)
    # COL = COL + 5
    # result2 = pd.DataFrame()
    # for arr_ in arr[2].values():
    #     pv1 = df[df['办卡地址省']==arr_]
    #     result2 = pv1.append(result2)
    # pv_result2 = pd.pivot_table(result2, index=['年', '办卡地址省'], values='就诊人次',aggfunc='sum')
    # pv_result2.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)
    writer.save()

