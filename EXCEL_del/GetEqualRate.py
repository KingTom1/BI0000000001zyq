import difflib

# 相似度
def GetEqualRate(str1,str2):
	return difflib.SequenceMatcher(None,str1,str2).quick_ratio()

# print(GetEqualRate("门诊挂号记录关联度患者","挂号表关联患者就诊基本信息"))


