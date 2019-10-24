import win32com.client
"""
	THIS MODULE IS THE WEAP-BACKEND FOR FEWSIM SYSTEM
"""
def get_WEAP_Results():
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	WEAP.ActiveArea = 'Internal_Linking_test_das'
	WEAP.ActiveScenario = WEAP.Scenarios[1]
	for b in WEAP.Branches:
		print(b.Name)

	print(len(WEAP.Branches))
	win32com.CoUninitialize()

get_WEAP_Results()