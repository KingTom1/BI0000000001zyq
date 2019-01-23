import xlrd
import pandas as pd
import EXCEL_del.GetEqualRate as Rate
import EXCEL_del.DB as DB


class ReadTwoExcel:
    def __init__(self, excel_path, excel_path1, sheet, sheet1, col_name1, col_sql):
        self.excel_path = excel_path
        self.excel_path1 = excel_path1
        self.sheet = sheet
        self.sheet1 = sheet1
        self.col_name1 = col_name1
        self.col_sql = col_sql

    # 读取表中第一行的值
    def read_cow(self):
        # excel_path = r"委预算\2018-11-23 (24号文件)XX医院-数据质量调整进度.xlsx"
        workbook = xlrd.open_workbook(self.excel_path)
        worksheet = workbook.sheet_by_name(self.sheet)
        num_cols = worksheet.ncols
        firstrow_data = []
        col = worksheet.col_values(1)
        for curr_col in range(len(col)):
            if len(col[curr_col]):
                firstrow_data.append(col[curr_col])
        return firstrow_data

    # pandas读取Excel
    def readExcel(self, epath, sname):
        return pd.read_excel(epath, sheet_name=sname)

    # 处理第一张表的关键字段
    def GetKeyCode(self, firstrow_data, d):
        # d = pd.read_excel(excel_path,sheet_name=u'医疗数据量及关联度')
        KeyCodes = []
        for i in range((len(firstrow_data) - 4)):
            CodeDesc = firstrow_data[i + 4]
            CodeGL = d[CodeDesc][0]
            KeyCode = CodeDesc + "" + CodeGL
            KeyCode = KeyCode.replace("度（", "").replace("）", "").replace("度(", "").replace(")", "").replace("（", "")
            KeyCodes.append(KeyCode)
        return KeyCodes

    # 读取表中第一列的值
    def read_col(self, key):
        # excel_path1 = r"委预算\各表数据指标抽检.xlsx"
        workbook1 = xlrd.open_workbook(self.excel_path1)
        worksheet1 = workbook1.sheet_by_name(self.sheet1)
        num_rows = worksheet1.nrows
        firstKeyCol_data = []
        for curr_row in range(num_rows):
            row = worksheet1.row_values(curr_row)
            if row[0].find(key) != -1:
                firstKeyCol_data.append(row[0])
        return firstKeyCol_data

    # 计算两个字符串的相识度，取最优值
    def zidian(self, KeyCodes):
        zidian = {}
        for Keycode in KeyCodes:
            index = Keycode.find("关联")
            key = Keycode[index:]
            if key == "关联评估报告":
                key = "关联病案"
            if key == "关联实验室记录":
                key = "关联实验室"
            firstKeyCol_data = self.read_col(key)
            # print(Keycode)
            t = 0
            for fc in firstKeyCol_data:
                rated1 = Rate.GetEqualRate(fc, Keycode)
                if t < rated1:
                    t = rated1
                    zidian[Keycode] = fc
            # print(t)
        return zidian

    # 处理字典，取值组成数组
    def arr(self, k_v):
        arr = []
        for v in k_v.values():
            arr.append(v)
        return arr

    # 处理字典，取K值组成数组
    def arr_k(self, k_v):
        arr = []
        for k in k_v.keys():
            arr.append(k)
        return arr

    # 处理pandas中取对应字段需要的行值
    def get_sql_strs(self,d1,arr):
        d2 = d1.set_index(d1[self.col_name1])
        new_d1 = d2.reindex(index=arr, columns=[self.col_sql])
        return new_d1

    # 根据sql查询出值
    def run_sql(self,sqls,zidian,year,cursor,arr_k,arr):
        df_arr = pd.DataFrame()
        for row in sqls.iteritems():
            for i in range(len(zidian)):
                sql = row[1][i]
                if sql.find('2016') != -1:
                    sql = sql.replace('2016', year)
                if sql.find('2017') != -1:
                    sql = sql.replace('2017', year)
                if sql.find('2018') != -1:
                    sql = sql.replace('2018', year)
                df = DB.get_result(cursor, sql)
                df['keyCode'] = arr_k[i]
                df[self.col_name1] = arr[i]
                df['年份'] = year
                print(df)
                df_arr = pd.concat([df_arr, df], axis=0, join='outer')
        return df_arr
