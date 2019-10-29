import win32com.client
"""
	THIS MODULE IS THE WEAP-BACKEND FOR FEWSIM SYSTEM
"""
def get_WEAP_Outputs():
	"""
	This function extract all results values from WEAP
	:return: Structured dictionary of WEAP results value
	"""
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	WEAP.ActiveArea = 'Internal_Linking_test_das'
	WEAP.ActiveScenario = WEAP.Scenarios[1]
	for b in WEAP.Branches:
		for v in b.Variables:
			print(b.FullName)
			if v.IsResultVariable == True:
				print(v.name)
				print(type(WEAP.Branch(b.FullName)))
				path = b.FullName + ":" +v.name
				print(WEAP.ResultValue(path, 2002, 1, 'Linkage', 2002,12, 'Total'))

	print(len(WEAP.Branches))
	win32com.CoUninitialize()

get_WEAP_Outputs()
# WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
#
# print(WEAP.Branch('\Demand Sites and Catchments\Municipal').Variables('Supply Requirement').Value)
# print(WEAP.ResultValue('\Demand Sites and Catchments\Municipal: Supply Requirement', 2002, 1, 'Linkage', 2002,12, 'Total'))

# for v in WEAP.Branch('Demand Sites and Catchments\Mun
# icipal').Variables:
# 	if v.IsResultVariable == True:
# 		print(v.Name, v.Value)