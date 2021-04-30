import win32com.client
import json
from os import listdir
import pandas as pd
import numpy as np
from datetime import datetime
import pythoncom


class Climate_Data():
    def __init__(self):
        self.root_path = "D:\\Project\\Food_Energy_Water\\fewsim-backend\\climate_data"

    def generate_default_variables(self):
        pythoncom.CoInitialize()
        WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
        climate_variables = ["Precipitation", "ETref", "Min Humidity", "Wind"]
        catchments = ["Maricopa_Water_District", "Roosevelt_ID", "Buckeye", "Tonopah", "Salt_River_Valley",
                      "Saint_Joins", "Roosevelt_Water_Consertation", "Queen_Creek", "Peninsula", "New_Magma",
                      "Arlington", "Adaman"]
        scenarios = ["Current Accounts", "Reference"]
        default_expression = []
        for site in catchments:
            for variable in climate_variables:
                for s in scenarios:
                    WEAP.ActiveScenario = s
                    default_expression.append(["Demand Sites and Catchments\\" + site, variable, s,
                                               WEAP.Branch("Demand Sites and Catchments\\" + site).Variable(
                                                   variable).Expression])
        print(default_expression)
        columns = ["branch", "variable", "scenario", "default_expression"]
        default_expression = pd.DataFrame(default_expression, columns=columns)
        default_expression.to_csv("climate_default_expression.csv")
        pythoncom.CoUninitialize()

    def intiate_climate_data(self):

        path = self.root_path
        CMIP5 = listdir(path + "\\CMIP5")
        CMIP6 = listdir(path + "\\CMIP6")
        climate_scenarios = {"CMIP5": [], "CMIP6": []}
        for file in CMIP5:
            climate_scenarios["CMIP5"].append(
                {"path": path + "\\CMIP5" + "\\" + file, "file_name": file})
        for file in CMIP6:
            climate_scenarios["CMIP6"].append(
                {"path": path + "\\CMIP6" + "\\" + file, "file_name": file})
        with open(path + "\\climate_files.json", "w") as outfile:
            json.dump(climate_scenarios, outfile)
        print(climate_scenarios)

    def set_climate_MABIA(self, scenario_name=None, scenatio_type=None):
        pythoncom.CoInitialize()
        xl = win32com.client.Dispatch("Excel.Application")
        WEAP_catchments = {"MaricopaWD": "Maricopa_Water_District", "RooseveltID": "Roosevelt_ID", "Buckeye": "Buckeye",
                           "Tonopah": "Tonopah", "SaltRiverVal": "Salt_River_Valley",
                           "Saintjohns": "Saint_Joins", "RooseveltWC": "Roosevelt_Water_Consertation",
                           "QueenCreek": "Queen_Creek", "Peninsula": "Peninsula", "NewMagma": "New_Magma",
                           "Arlington": "Arlington", "Adaman": "Adaman"}
        WEAP_scenarios = ["Current Accounts", "Reference"]
        WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
        climate_variables = ["Precipitation", "ETref", "Min Humidity", "Wind"]
        with open(self.root_path + "\\climate_files.json", "r") as files:
            climate_files = json.load(files)
            for W_scenario in WEAP_scenarios:
                WEAP.ActiveScenario = W_scenario
                for k in climate_files.keys():
                    for k_f in climate_files[k]:
                        name = k_f["file_name"].split("_")[0]
                        catchment = k_f["file_name"].split("_")[1]
                        CMIP_type = k_f["file_name"].split("_")[2]
                        s_type = k_f["file_name"].split("_")[3].split(".")[0]
                        expression = "ReadFromFile({}, {}, , , , , , , , Cycle)"
                        if name == scenario_name:
                            if s_type == scenatio_type:
                                for variable in climate_variables:
                                    WEAP.Branch("Demand Sites and Catchments\\" + WEAP_catchments[catchment]).Variable(
                                        variable).Expression = expression.format(k_f["path"],
                                                                                 climate_variables.index(variable) + 1)
                                    xl.Quit()
                                    print("Set ", W_scenario, " ", "Demand Sites and Catchments\\" + WEAP_catchments[catchment],
                                          " to climate value: ", expression.format(k_f["path"], climate_variables.index(variable) + 1), " ", variable )

        pythoncom.CoUninitialize()

    def set_climate_default(self):
        pythoncom.CoInitialize()
        xl = win32com.client.Dispatch("Excel.Application")
        WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
        WEAP_scenarios = ["Current Accounts", "Reference"]
        default_expression = pd.read_csv(self.root_path + "\\climate_default_expression.csv", index_col=0)
        for W_scenario in WEAP_scenarios:
            WEAP.ActiveScenario = W_scenario
            for i in range(len(default_expression)):
                if default_expression.loc[i, "scenario"] == W_scenario:
                    WEAP.Branch(default_expression.loc[i, "branch"]).Variable(
                        default_expression.loc[i, "variable"]).Expression = default_expression.loc[
                        i, "default_expression"]
                    xl.Quit()
                    print("Set ", W_scenario, " ", default_expression.loc[i, "branch"], " back to default value: ",
                          default_expression.loc[i, "default_expression"], " ", default_expression.loc[i, "variable"])
        pythoncom.CoUninitialize()



# Climate_Data().set_climate_MABIA(scenario_name="CanESM5", scenatio_type="ssp585")
# Climate_Data().generate_default_variables()
# Climate_Data().set_climate_default()
# print('MPI-ESM1-2-HR_Adaman_CMIP6_Hist.csv'.split("_"))
# climate_variables = ["Precipitation", "ETref", "Min Humidity", "Wind"]
# print(climate_variables.index("ETref"))
