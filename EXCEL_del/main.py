import pandas as pd
import EXCEL_del.read_excle as read_excle
if __name__ == '__main__':
    # 自定义路径
    excel_path = r"excels\规则.xlsx"
    excel_path1 = r"excels\病人来源201810-12.xlsx"
    sheets = ['成都市','四川省异地','外省']
    sheet1 = u'门诊201810-12'
    j = 3
    col = 0
    row = 2
    writer = pd.ExcelWriter('统计结果表.xlsx')
    # 获取数据源数据
    df = pd.read_excel(excel_path1, sheet_name=sheet1)
    total = sum(df['就诊人次'])
    v_total = dict()
    for sheet in sheets:
        # 获取类对象
        RE = read_excle.ReadTwoExcel(excel_path, sheet)
        # 获取'规则表'中的数据
        guizeData = RE.read_cow(excel_path, sheet, 0)
        # 获取'数据源表中'的规则对应的数据,形成字典
        ODate = RE.read_cow(excel_path1, sheet1, j)
        zidian = dict()
        for a_ in guizeData:
            for b_ in ODate:
                if b_.find(a_)>0:
                    zidian[a_] = b_
        pv1 = pd.pivot_table(df,aggfunc='sum',values='就诊人次',index=guizeData[0])
        d = dict()
        for k in zidian.keys():
            d[k] = pv1.ix[zidian[k]]
        result = pd.DataFrame.from_dict(d,orient='index')
        result = result.sort_values('就诊人次',ascending=False)
        v_total[sheet]=sum(result['就诊人次'])
        # 各行数据总和求和并新增一行
        result.loc['合计'] = result.apply(lambda x: x.sum())
        result_ = result.reset_index()
        result_.insert(0, '类别', sheet)
        result_.to_excel(writer, sheet_name=sheet1, index=False, startrow=row, startcol=col)
        col = col+4
        j = j-1
    a = "%s四川省门诊病人来源地汇总"%sheet1
    b = "根据信息中心BI提供的数据，我院门诊就诊人次达%s人次，具体来源地如下:"%total
    heard = pd.DataFrame({a: [b]})
    heard.to_excel(writer,sheet_name=sheet1, index=False, startrow=0, startcol=0)
    v_total['不详/空']=total-v_total['成都市']-v_total['四川省异地']-v_total['外省']
    v_total['合计'] = total
    vv_total = pd.DataFrame.from_dict(v_total, orient='index')
    vv_total.to_excel(writer,sheet_name=sheet1, startrow=row, startcol=12)
    writer.save()

# # 横向拼接DateFrame数据
# results = pd.concat(results, axis=1, join_axes=[results[2].index])
#
