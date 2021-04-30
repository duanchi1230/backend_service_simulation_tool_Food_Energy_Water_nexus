"""
This scrip provides the API for the Agricultural model to be connected to the WEAP-MABIA module
Author: Chi Duan
Late Updated: 8/10/2020
"""
import numpy as np
import pandas as pd
import win32com.client
import pythoncom
import MPMmodel.StatsModel_2_5 as MPMmodel
import math
import os
from os import listdir
import json

class MPM():

    def __init__(self):
        self.root_path = "D:\Project\Food_Energy_Water\\fewsim-backend"
        self.dummy = 1
        pathTfile = self.root_path+'/MPMmodel/Averagefinal.csv'
        pathPfile = self.root_path+'/MPMmodel/FuturedataN.csv'
        out = self.root_path+'/MPMmodel/outPuts.csv'
        #
        """
          This (below) is the path to your R program and version numbe
        """
        Rpath = 'C:/Program Files/R/R-3.6.2/bin/R'
        """
         Chi;
          YOU WILL NEED to hard code a temporary directory and file name for
          a temporary csv file that the R code uses. I could not make it
          generic; R did not like it. lines 111, 113, 116, and 118
        """
        # tempPath='C:/data/R/out_temp.csv'
        # Write an output file of the results - True or False
        writeOut = True
        # Call the program using Main(var1,var2)
        # Main(pathTfile,pathPfile,writeOut,out)
        MPMmodel.Main(pathTfile, pathPfile, Rpath, writeOut, out)
        # ============================================================
        # E.O.F.
        self.Catchment_Variables_default = self.transform_to_input()

    def transform_to_input(self, climate_input=False):
        WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
        start_year = WEAP.BaseYear
        end_year = WEAP.EndYear
        varlist = ['Cotton', 'Alfalfa', 'Corn', 'Barley', 'Durum', 'Vegetables(potatoes)', 'Remaining']
        if climate_input == False:
            outPut = pd.read_csv(self.root_path+"/MPMmodel/outPuts.csv")
        else:
            outPut = pd.read_csv(climate_input)
            outPut.columns = ["Year", "0", "1", "2", "3", "4", "5", "6"]
            outPut.to_csv(self.root_path+"/MPMmodel/outPuts.csv", index=False)

        Catchment_Variables = pd.read_csv(self.root_path+"/MPMmodel/W_variables.csv", index_col=0)
        start_year = 2008
        end_year = 2050
        time_series_input = {}

        for v in varlist:
            time_series_input[v] = {}
        for i in range(len(outPut)):
            if outPut.loc[i, "Year"] >= start_year and outPut.loc[i, "Year"] <= end_year:
                for j in range(len(varlist)):
                    time_series_input[varlist[j]][str(outPut.loc[i, "Year"])] = outPut.loc[i, outPut.columns[j+1]]
        # print(time_series_input)
        for i in range(len(Catchment_Variables)):
            # time_series = Catchment_Variables.loc[i, "proportion"] * np.array(time_series_input[Catchment_Variables.loc[i, "crop_type"]])
            input_string = "Interp("
            for k in time_series_input[Catchment_Variables.loc[i, "crop_type"]].keys():
                input_string = input_string + k + "," + str(
                    time_series_input[Catchment_Variables.loc[i, "crop_type"]][k] * Catchment_Variables.loc[
                        i, "proportion"] * 100) + ", "
            input_string = input_string[:-2] + ")"
            Catchment_Variables.loc[i, "time_series"] = input_string

            self.Catchment_Variables = Catchment_Variables

        return Catchment_Variables

    def set_MPM_percentage(self, WEAP, percentage, branchToChange, variableToChange):
        FuturedataN = pd.read_csv(self.root_path+"/MPMmodel/FuturedataN.csv")
        Averagefinal = pd.read_csv(self.root_path+"/MPMmodel/Averagefinal.csv")
        y_crop = ["cotton", "corn", "barley", "durum", "alfalfa", "vegetables"]
        p_crop = ["cotton", "corn", "barley", "durum", "alfalfa"]
        if variableToChange == "yield":
            c = branchToChange.split("\\")[-1]
            # print(c, percentage)
            FuturedataN["y_" + c] = FuturedataN["y_" + c] * float(percentage)
            # for c in y_crop:
            #     FuturedataN["y_" + c] = FuturedataN["y_" + c] * percentage
            #     Averagefinal["y_" + c] = Averagefinal["y_" + c] * percentage
        if variableToChange == "price":
            c = branchToChange.split("\\")[-1]
            FuturedataN["p_" + c] = FuturedataN["p_" + c] * float(percentage)
            # Averagefinal["p_" + c] = Averagefinal["p_" + c] * percentage
            # .iloc[[0, 1, 2]]
        FuturedataN.to_csv(self.root_path+"/MPMmodel/FuturedataN_interpreter.csv")
        Averagefinal.to_csv(self.root_path+"/MPMmodel/Averagefinal_interpreter.csv")

        pathTfile = self.root_path+'/MPMmodel/Averagefinal_interpreter.csv'
        pathPfile = self.root_path+'/MPMmodel/FuturedataN_interpreter.csv'
        out = self.root_path+'/MPMmodel/outPuts.csv'
        Rpath = 'C:/Program Files/R/R-3.6.2/bin/R'
        writeOut = True
        MPMmodel.Main(pathTfile, pathPfile, Rpath, writeOut, out)
        self.transform_to_input()
        self.set_MPM_MABIA(WEAP)

    def set_MPM_climate(self, WEAP, climate_input):
        climate_input = 'D:\\Project\\Food_Energy_Water\\fewsim-backend/MPMmodel//climate/CMIP 5 models/RCP 45/ACCESS1_0_PR_MRC.csv'
        self.transform_to_input(climate_input=climate_input)
        self.set_MPM_MABIA(WEAP)

    def set_MPM_MABIA(self, WEAP):

        for s in WEAP.Scenarios:
            WEAP.ActiveScenario = s
            for i in range(len(self.Catchment_Variables)):
                WEAP.Branch("Demand Sites and Catchments\\" + self.Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                            self.Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression = \
                    self.Catchment_Variables.iloc[i]["time_series"]

            # print(self.Catchment_Variables.iloc[i])

            crop_area = pd.read_csv(self.root_path+"/MPMmodel/totalCropArea.csv")
            for i in range(len(crop_area)):
                WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["averageArea"]

    def set_MPM_default(self, WEAP):
        xl = win32com.client.Dispatch("Excel.Application")
        for i in range(len(self.Catchment_Variables_default)):
            WEAP.Branch(
                "Demand Sites and Catchments\\" + self.Catchment_Variables_default.iloc[i]["demand_site"] + "\\" +
                self.Catchment_Variables_default.iloc[i]["crop"]).Variable("Area").Expression = \
                self.Catchment_Variables_default.iloc[i]["time_series"]
            xl.Quit()
            # print(self.Catchment_Variables_default.iloc[i])
        crop_area = pd.read_csv(self.root_path+"/MPMmodel/totalCropArea.csv")
        for i in range(len(crop_area)):
            WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["averageArea"]
            xl.Quit()

    def decouple_MPM_MABIA(self, WEAP):
        xl = win32com.client.Dispatch("Excel.Application")
        for i in range(len(self.Catchment_Variables)):
            WEAP.Branch("Demand Sites and Catchments\\" + self.Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                        self.Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression = \
                self.Catchment_Variables.iloc[i]["default_expression"]
            xl.Quit()
            # print(self.Catchment_Variables.iloc[i])
        crop_area = pd.read_csv(self.root_path+"/MPMmodel/totalCropArea.csv")
        for i in range(len(crop_area)):
            WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["default_expression"]
            xl.Quit()

    def get_default_expression(self):
        Catchment_Variables = pd.read_csv("W_variables.csv", index_col=0)
        WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
        for i in range(len(Catchment_Variables)):
            Catchment_Variables.loc[i, "default_expression"] = str(WEAP.Branch(
                "Demand Sites and Catchments\\" + Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression)
            print(WEAP.Branch(
                "Demand Sites and Catchments\\" + Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression)
        # Catchment_Variables.loc[0, "default_expression"] = 1
        Catchment_Variables.to_csv("W_variables.csv")

    def get_default_area(self):
        crop_area = pd.read_csv(self.root_path+"/MPMmodel/totalCropArea.csv")
        WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
        start = WEAP.BaseYear
        end = WEAP.EndYear
        # print(crop_area.iloc[0])
        total = 0
        for i in range(len(crop_area)):
            value = []
            for y in range(WEAP.BaseYear, WEAP.EndYear + 1):
                print(crop_area.iloc[i]["branch"] + crop_area.iloc[i]["variable"] + ":[M^2]")
                value.append( WEAP.ResultValue(
                    crop_area.iloc[i]["branch"] + ":" + crop_area.iloc[i]["variable"] + "[M^2]", y, 1, "Reference", y,
                    WEAP.NumTimeSteps,
                    "Average"))
            total = total + sum(value)/len(value)
            crop_area.loc[i, "default_expression"] = WEAP.Branch(crop_area.iloc[i]["branch"]).Variable(crop_area.iloc[i]["variable"]).Expression

        for i in range(len(crop_area)):
            crop_area.loc[i, "averageArea"] = total
        crop_area.to_csv(self.root_path+"/MPMmodel/totalCropArea.csv")
    # def find_climate_files(self):
    #     os.

    def intialize_climate_files(self):
        root_path = self.root_path
        path = root_path+"/MPMmodel/"
        CMIP5 = listdir("./climate/CMIP 5 models")
        CMIP6 = listdir("./climate/CMIP 6 models")
        climate_scenarios = {"CMIP5":{}, "CMIP6":{}}
        for folder in CMIP5:
            climate_scenarios["CMIP5"][folder] = []
            for file in listdir(path+"/climate/CMIP 5 models/"+folder):
                climate_scenarios["CMIP5"][folder].append({"path":path+"/climate/CMIP 5 models/"+folder+"/"+file, "file_name":file})
        for folder in CMIP6:
            climate_scenarios["CMIP6"][folder] = []
            for file in listdir(path+"/climate/CMIP 6 models/"+folder):
                climate_scenarios["CMIP6"][folder].append({"path":path+"/climate/CMIP 6 models/"+folder+"/"+file, "file_name":file})
        with open(root_path+"/MPMmodel/climate/"+"climate_files.json", "w") as outfile:
            json.dump(climate_scenarios, outfile)
        print(climate_scenarios)


# The following is some testing script
# WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
# MPM = MPM()
# MPM.set_MPM_default(WEAP)
# MPM.set_MPM_climate(WEAP, climate_input="D:\\Project\\Food_Energy_Water\\fewsim-backend/MPMmodel//climate/CMIP 5 models/RCP 45/ACCESS1_0_PR_MRC.csv")
# MPM.decouple_MPM_MABIA(WEAP)
# data = pd.read_csv("D:\\Project\\Food_Energy_Water\\fewsim-backend/MPMmodel//climate/CMIP 5 models/RCP 45/ACCESS1_0_PR_MRC.csv")
# print(data)
# data.columns=["Year", "0", "1", "2", "3", "4", "5", "6"]
# data.to_csv("asd.csv", index=False)