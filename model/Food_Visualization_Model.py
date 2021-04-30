import pandas as pd
import numpy as np
import pythoncom
import win32com.client


def get_WEAP_value(branch, variable, type=None):
    WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
    start_year = WEAP.BaseYear
    end_year = WEAP.EndYear
    value_year = []

    for year in range(start_year + 1, end_year + 1):
        value_year.append(round(WEAP.ResultValue(branch + ":" + variable, year, 1, "Reference", year, 12, 'Average'), 1))
    return value_year


def get_food_variables():
    WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
    start_year = WEAP.BaseYear
    end_year = WEAP.EndYear
    root_path = "D:\Project\Food_Energy_Water\\fewsim-backend"
    WEAP_variabels = pd.read_csv(root_path + "\model\W_variables.csv", index_col=0)
    food_variables = WEAP_variabels.loc[(WEAP_variabels["variable-name"] == "Annual Crop Production") | (
                WEAP_variabels["variable-name"] == "Area Calculated")].copy()
    food_variables = food_variables.set_index(pd.Index(list(range(len(food_variables)))))
    years = list(range(start_year + 1, end_year + 1))
    food_variables[years] = 0
    food_result = {"weap":[], "mpm":[]}
    for i in list(food_variables.index):
        food_variables.loc[i, years] = get_WEAP_value(food_variables.loc[i, "branch"],
                                                      food_variables.loc[i, "variable-name"])
        food_result["weap"].append(
            {"branch": food_variables.loc[i, "branch"], "variable": food_variables.loc[i, "variable-name"],
             "value": get_WEAP_value(food_variables.loc[i, "branch"], food_variables.loc[i, "variable-name"])})
    crops = ['cotton', 'alfalfa', 'corn', 'barley', 'durum', 'veg', 'remaining']
    total_Croprea = 439100836.4
    mpm_outputs = pd.read_csv(root_path + "\MPMmodel\outPuts.csv", index_col=0)
    for col in mpm_outputs.columns:
        food_result["mpm"].append({"crop": crops[int(col)], "value": (mpm_outputs.loc[start_year+1:end_year, col].to_numpy()*total_Croprea).tolist()})
        # print(mpm_outputs.loc[start_year+1:end_year, col].to_numpy().tolist())
    print(food_result)
    return food_result

def get_mpm_variables_tree():
    pythoncom.CoInitialize()
    WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
    start_year = WEAP.BaseYear
    end_year = WEAP.EndYear
    root_path = "D:\Project\Food_Energy_Water\\fewsim-backend"
    crops = ['cotton', 'alfalfa', 'corn', 'barley', 'durum', 'veg', 'remaining']
    food_output = {"name": "food-output", "children":[]}
    total_Croprea = 439100836.4
    mpm_model_outputs = pd.read_csv(root_path + "\MPMmodel\outPuts.csv", index_col=0)
    food_inputs = {"name": "food-input", "children": []}
    for col in mpm_model_outputs.columns:
        food_output["children"].append({"name": crops[int(col)], "model": "mpm", "fullname": "Total_Area", "value": (
                    mpm_model_outputs.loc[start_year + 1:end_year, col].to_numpy() * total_Croprea).tolist()})
    pythoncom.CoUninitialize()
    return food_inputs, food_output
# get_food_variables()
# WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
#
# print(WEAP.ResultValue('Demand Sites and Catchments\\New_Magma\\Alfalfa_hay:Area Calculated', 2019, 1, "Reference",
# 					                  2019, 12, 'Average'))
# print(WEAP.ResultValue('Demand Sites and Catchments\\New_Magma\\Alfalfa_hay:Area', 2019, 1, "Reference",
# 					                  2019, 12, 'Average'))
