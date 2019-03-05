import pandas as pd
import math
import EXCEL_del.read_excle as read_excle

excel_path1 = r"excels\病人来源统计20190228.xlsx"
sheet1 = u'Sheet1'



df = pd.read_excel(excel_path1, sheet_name=sheet1)
df1 = pd.read_excel(excel_path1, sheet_name=sheet1,index_col='月份')
arr = list(set(df['月份']))
for arr_ in arr:
    # print(arr_)
    arr_df = df1.ix[arr_]
    pivot = pd.pivot_table(arr_df,index=["办卡地址省","办卡地址市","办卡地址区县"],aggfunc='sum',values='就诊人次')
    print(pivot)
    # print(pivot.ix['武侯区'])
    # arr_df_ =list(arr_df.iloc[:,3])
    # print(arr_df_[1]!=arr_df_[1])

