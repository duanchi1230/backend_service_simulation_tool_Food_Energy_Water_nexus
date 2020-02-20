import win32com.client
import json
import pandas as pd
import numpy as np
import pythoncom
### This is the new path on your machine ###
new_drive_path = "D:\\Project\\Food_Energy_Water\\Data_Used_WEAP"

### This is the old path from the migrated model ###
old_dirve_path = "C:\\Users\\amounir\Dropbox (ASU)\\344\\NEW WEAP-MABIA Model"
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
start_year = WEAP.BaseYear
end_year = WEAP.EndYear
WEAP.ActiveArea = 'Ag_MABIA_v14'
for s in LEAP.Scenarios:
	if s != 'Current Account':
		active_scenario = s
LEAP.ActiveScenario = active_scenario
print(LEAP.ActiveScenario)

### The following updates the old path with the new path for all (existing) scenarios in WEAP ###
for scenario in WEAP.Scenarios:
	WEAP.ActiveScenario = scenario
	for b in WEAP.Branches:
		# WEAP.Branch(b.FullName)
		# print('\n')
		print(b.FullName)
		# print(b.Name)
		for v in WEAP.Branch(b.FullName).Variables:
			variable_list = pd.read_excel('WEAP_Input_Variables.xlsx')
			variable_list = np.array(variable_list['variable_name'])
			if v.name in variable_list:
				# print(v.name)
				if old_dirve_path in WEAP.Branch(b.FullName).Variable(v.name).Expression:
					WEAP.Branch(b.FullName).Variable(v.name).Expression = WEAP.Branch(b.FullName).Variable(v.name).Expression.replace(old_dirve_path, new_drive_path)
					print(WEAP.Branch(b.FullName).Variable(v.name).Expression.replace(old_dirve_path, new_drive_path))
### For current LEAP model, there are no paths for variables. ###