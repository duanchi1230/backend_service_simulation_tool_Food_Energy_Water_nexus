"""
This scrip provides the API for the Agricultural model to be connected to the WEAP-MABIA module
Author: Chi Duan
Late Updated: 8/10/2020
"""
import numpy as np
import pandas as pd
import win32com.client
import pythoncom
import StatsModel_2_5 as MPMmodel
import math

class MPM():

    def __init__(self):
        self.dummy = 1
        pathTfile = 'Averagefinal.csv'
        pathPfile = 'FuturedataN.csv'
        out = 'outPuts.csv'
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
        # self.Catchment_Variables_default = self.transform_to_input()

    def transform_to_input(self):
        varlist = ['Cotton', 'Alfalfa', 'Corn', 'Barley', 'Durum', 'Vegetables(potatoes)', 'Remaining']
        outPut = pd.read_csv("./MPMmodel/outPuts.csv")
        Catchment_Variables = pd.read_csv("./MPMmodel/W_variables.csv", index_col=0)
        start_year = 2008
        end_year = 2018
        time_series_input = {}

        for v in varlist:
            time_series_input[v] = {}
        for i in range(len(outPut)):
            if outPut.loc[i, "Year"] >= start_year and outPut.loc[i, "Year"] <= end_year:
                for j in range(len(varlist)):
                    time_series_input[varlist[j]][str(outPut.loc[i, "Year"])] = outPut.loc[i, str(j)]
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

    def set_MPM_percentage(self, percentage, variableToChange):
        FuturedataN = pd.read_csv("FuturedataN.csv")
        Averagefinal = pd.read_csv("Averagefinal.csv")
        y_crop = ["cotton", "corn", "barley", "durum", "alfalfa"]
        p_crop = ["cotton", "corn", "barley", "durum", "alfalfa"]
        if variableToChange == "yield":
            for c in y_crop:
                FuturedataN["y_" + c] = FuturedataN["y_" + c] * percentage
                # Averagefinal["y_" + c] = Averagefinal["y_" + c] * percentage
        if variableToChange == "price":
            for c in p_crop:
                FuturedataN["p_" + c] = FuturedataN["p_" + c] * percentage
                # Averagefinal["p_" + c] = Averagefinal["p_" + c] * percentage
                # .iloc[[0, 1, 2]]
        FuturedataN.to_csv("FuturedataN_interpreter.csv")
        Averagefinal.to_csv("Averagefinal_interpreter.csv")

        pathTfile = 'Averagefinal_interpreter.csv'
        pathPfile = 'FuturedataN_interpreter.csv'
        out = 'outPuts.csv'
        Rpath = 'C:/Program Files/R/R-3.6.2/bin/R'
        writeOut = True
        MPMmodel.Main(pathTfile, pathPfile, Rpath, writeOut, out)
        # self.transform_to_input()
        # self.set_MPM_MABIA(WEAP)

    def set_MPM_MABIA(self, WEAP):
        for i in range(len(self.Catchment_Variables)):
            WEAP.Branch("Demand Sites and Catchments\\" + self.Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                        self.Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression = \
                self.Catchment_Variables.iloc[i]["time_series"]
            # print(self.Catchment_Variables.iloc[i])

        crop_area = pd.read_csv("./MPMmodel/totalCropArea.csv")
        for i in range(len(crop_area)):
            WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["averageArea"]

    def set_MPM_default(self, WEAP):
        for i in range(len(self.Catchment_Variables_default)):
            WEAP.Branch(
                "Demand Sites and Catchments\\" + self.Catchment_Variables_default.iloc[i]["demand_site"] + "\\" +
                self.Catchment_Variables_default.iloc[i]["crop"]).Variable("Area").Expression = \
                self.Catchment_Variables_default.iloc[i]["time_series"]
            # print(self.Catchment_Variables_default.iloc[i])
        crop_area = pd.read_csv("./MPMmodel/totalCropArea.csv")
        for i in range(len(crop_area)):
            WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["averageArea"]

    def decouple_MPM_MABIA(self, WEAP):
        for i in range(len(self.Catchment_Variables)):
            WEAP.Branch("Demand Sites and Catchments\\" + self.Catchment_Variables.iloc[i]["demand_site"] + "\\" +
                        self.Catchment_Variables.iloc[i]["crop"]).Variable("Area").Expression = \
                self.Catchment_Variables.iloc[i]["default_expression"]
            # print(self.Catchment_Variables.iloc[i])
        crop_area = pd.read_csv("./MPMmodel/totalCropArea.csv")
        for i in range(len(crop_area)):
            WEAP.Branch(crop_area.iloc[i]["branch"]).Variables(crop_area.iloc[i]["variable"]).Expression = crop_area.iloc[i]["default_expression"]

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
        crop_area = pd.read_csv("./MPMmodel/totalCropArea.csv")
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
        crop_area.to_csv("./MPMmodel/totalCropArea.csv")

MPM = MPM()
# MPM.get_default_area()

input_variables = pd.read_csv("variables_input_table.csv", index_col=0)
print(len(input_variables))
values = []
for i in range(len(input_variables)):
    if input_variables.iloc[i]["type"]=="MPM":
        print(i)
        variables = ["yield", "price"]
        MPM.set_MPM_percentage(input_variables.iloc[i]["value"]/100, variables[input_variables.iloc[i]["i"]])
        outPuts = pd.read_csv(".outPuts.csv")
        for j in range(len(outPuts)):
            values_year = outPuts.iloc[j].tolist()
            values_year[0] = int(values_year[0])
            print(values_year)
            values_year.append(input_variables.iloc[i]["i"])
            values_year.append(input_variables.iloc[i]["value"])
            values_year.append(input_variables.iloc[i]["loop"])
            values_year.append(input_variables.iloc[i]["type"])
            values.append(values_year)
columns = ["year", "cotton", "alfalfa", "corn", "barley", "durum", "veg", "remaining", "i", "value", "loop", "type"]
pd_values = pd.DataFrame(values, columns=columns)
print(pd_values)
pd_values.to_csv("intermediate_crop_area.csv")