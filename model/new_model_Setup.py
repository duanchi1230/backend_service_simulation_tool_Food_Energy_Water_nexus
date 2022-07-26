import win32com.client
import json
import pandas as pd
import numpy as np
import pythoncom

### Excel Close Controller ###
xl = win32com.client.Dispatch("Excel.Application")


def set_model_path(new_drive_path, old_dirve_path):
    LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
    WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
    start_year = WEAP.BaseYear
    end_year = WEAP.EndYear
    # WEAP.ActiveArea = 'Ag_MABIA_v15'
    for s in LEAP.Scenarios:
        if s != 'Current Account':
            active_scenario = s
    # LEAP.ActiveScenario = active_scenario
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
                        WEAP.Branch(b.FullName).Variable(v.name).Expression = WEAP.Branch(b.FullName).Variable(
                            v.name).Expression.replace(old_dirve_path, new_drive_path)
                        xl.Quit()
                        print(WEAP.Branch(b.FullName).Variable(v.name).Expression.replace(old_dirve_path, new_drive_path))
### This is the new path on your machine ###
new_drive_path = "D:\\Project\\Food_Energy_Water\\Data_Used_WEAP\\Ag_MABIA_v14_50\\data"

### This is the old path from the migrated model ###
old_dirve_path = "D:\\Project\\Food_Energy_Water\\Data_Used_WEAP\\data"
set_model_path(new_drive_path=new_drive_path, old_dirve_path=old_dirve_path)


# ### For current LEAP model, there are no paths for variables. ###
# LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
# # LEAP.Branch('Key\\Population').Variable('Population').Expression = LEAP.Branch('Key\\Population').Variable('Population').Expression +str(1)
# print(LEAP.Branch('Key Assumptions\\Population').Variable('Population').Expression)
