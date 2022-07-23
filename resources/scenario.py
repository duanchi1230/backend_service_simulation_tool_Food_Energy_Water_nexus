from flask import Flask, request
from flask_restful import Resource, Api, abort
from model import WEAP_Visualization_Model as WEAP_model
from sensitivity_index import get_data
from model import LEAP_Visualization_Model as LEAP_model
from model import Food_Visualization_Model as food_model
import win32com.client
from model import weap_backend, leap_backend, run_scenarios
from model import run_scenarios
from model.database import Scenarios, session
import json
from datetime import datetime
import os
import pandas as pd


class Scenario(Resource):

    def __init__(self):
        pass

    def get(self, pid, model, sid):
        sid = int(sid)

        # Chi: implement a function that gets a scenario based on the "model" and the "sid"
        # The returned dict includes all the input and output variables

        flow, timeRange = WEAP_model.get_WEAP_flow_value()
        value = flow[list(flow.keys())[sid]]

        return {
                   "sid": sid,
                   "name": list(flow.keys())[sid],
                   "runStatus": "finished",
                   "timeRange": timeRange,
                   "numTimeSteps": timeRange[1] - timeRange[0],
                   "var": {"input": "1230",
                           "output": value}
               }, 200

    def post(self, pid, model, sid):
        data = request.get_json(self)

        win32com.CoInitialize()
        WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
        start_year = WEAP.BaseYear
        end_year = WEAP.EndYear
        WEAP.Calculate()
        win32com.CoUninitialize()
        return {"response": "POST method was just called!"}, 201


class ScenarioList(Resource):

    def __init__(self):
        pass

    def get(self, pid, model):
        # Chi: implement a function that gets all the existing scenarios
        # and return their brief summary as follows:
        scenario_list = []
        scenarios = ["Reference", "5% Population Growth", "10% Population Growth"]
        for i in range(3):
            # get the object of the scenario here

            year = [1986, 2008]

            value = WEAP_model.get_WEAP_flow_value()
            scenario_summary = {
                "sid": i,
                "name": scenarios[i],
                "runStatus": "finished",
                "timeRange": year,
                "numTimeSteps": (year[1] - year[0]),
                "__filled": False
            }

            scenario_list.append(scenario_summary)

        return scenario_list, 200

    # return [{
    #     "sid": "name of the scenario",
    #     "runStatus": "finished",
    #     "timeRange": ["range start", "range end"],
    #     "numTimeSteps": "number of time steps",
    # }, {
    #     "sid": "name of the scenario",
    #     "runStatus": "finished",
    #     "timeRange": ["range start", "range end"],
    #     "numTimeSteps": "number of time steps"
    # }], 200

    def post(self, pid, model):
        """
        Create a new scenario based on a reference
        :param pid:
        :param model:
        :return:
        """
        # return the newly-created scenario together with CREATED 201 status code
        return {"a": model}, 201


class Input_List(Resource):

    def __init__(self):
        pass

    def get(self, format):

        if format == "tree":
            weap_variables = weap_backend.get_WEAP_variables_tree("./model/WEAP_variables.json")
            leap_variables = leap_backend.get_LEAP_variables_tree("./model/LEAP_variables.json")
            food_inputs, food_output = food_model.get_mpm_variables_tree()
        if format == "flat":
            weap_variables = weap_backend.get_WEAP_variables_from_file("./model/WEAP_variables.json")
            leap_variables = leap_backend.get_LEAP_variables_from_file("./model/LEAP_variables.json")
            food_inputs, food_output = food_model.get_mpm_variables_tree()
        data = {"name": "FEW Nexus-Variables",
                "model": "FEW",
                "type": "none",
                "children": [weap_variables[0], weap_variables[1], leap_variables[0], leap_variables[1], food_inputs, food_output],
                }
        root_path = "D:\\Project\\Food_Energy_Water\\fewsim-backend\\model"
        climate_file = pd.read_csv(root_path + "\\Climate_variables.csv", index_col=0)
        climate_scenarios = []
        for i in range(len(climate_file)):
            climate_scenarios.append({"name": climate_file.loc[i, "name"], "type": climate_file.loc[i, "type"],
                                      "CMIP": climate_file.loc[i, "CMIP"]})
        return {"data":data, "climate_scenarios": climate_scenarios}, 200

    def post(self):
        return {"response": "POST method was just called!"}, 201


class Run_Sceanrios(Resource):
    def __init__(self):
        pass

    def get(self, scenario):
        with open("run_results.json", "r") as outfile:
            result = json.load(outfile)
        now = datetime.now()
        dt_string = now.strftime('%d/%m/%Y %H:%M:%S')
        log = pd.DataFrame([[dt_string, 'Started']], columns=['time', 'message'])
        log.to_csv('log.csv')
        return result

    def post(self, scenario):
        # print(scenario)
        packed_scenarios = request.get_json(self)
        scenarios = packed_scenarios[0]
        sustainability_variables = packed_scenarios[1]
        loaded_group_index = packed_scenarios[2]
        with open("para.json", "w") as outfile:
            json.dump(packed_scenarios, outfile)
        weap_flow, leap_data, food_data, s_variables, loaded_group_index = run_scenarios.run_all_secanrios(scenarios,
                                                                                                sustainability_variables,
                                                                                                loaded_group_index)
        # print(weap_flow, leap_data)
        return {"weap-flow": weap_flow, "leap-data": leap_data, "food-data":food_data, "sustainability_variables": s_variables,
                "loaded_index_group": loaded_group_index}, 201


class Get_Run_Log(Resource):
    def __init__(self):
        pass

    def get(self):
        run_log_file = pd.read_csv("log.csv")
        log = []
        # print(run_log_file)
        for row in run_log_file.iterrows():
            # print({"time": str(row[1]["time"]), "message": str(row[1]["message"])})
            log.append({"time": str(row[1]["time"]), "message": str(row[1]["message"])})
        return log, 200

    def post(self, scenario):
        pass


class Load_Existing_Scenarios(Resource):
    def __init__(self):
        pass

    def get(self):
        with open(".\\scenarios\\existing_scenarios.json", "r") as file:
            scenarios = json.load(file)
        return scenarios, 200

    def post(self):
        scenarios = request.get_json(self)
        with open(".\\scenarios\\existing_scenarios.json", "w") as file:
            json.dump(scenarios, file)
        return "Scenarios have been updated!", 201


class Get_Coupled_Variable(Resource):
    def __init__(self):
        pass

    def get(self):
        df = pd.read_csv(".\\model\\coupled_parameters.csv")
        print(df.iloc[0])
        coupled_parameters = df.to_numpy().tolist()
        return coupled_parameters, 200

    def post(self, scenario):
        pass


class Save_Simulation_Result(Resource):
    def __init__(self):
        pass

    def get(self):
        with open(".\\simulation\\simulation_result.json", "r") as file:
            simulation_result = json.load(file)
        return simulation_result, 200

    def post(self):
        simulation_result = request.get_json(self)
        with open(".\\simulation\\simulation_result.json", "w") as file:
            json.dump(simulation_result, file)
        return "Simulation Result is Saved Successfully!", 201


class Load_Sustainability_Index(Resource):
    def __init__(self):
        pass

    def get(self):
        with open(".\\sustainability_index\\sustainability_index.json", "r") as file:
            sustainability_index = json.load(file)
        return sustainability_index, 200

    def post(self):
        sustainability_index = request.get_json(self)
        with open(".\\sustainability_index\\sustainability_index.json", "w") as file:
            json.dump(sustainability_index, file)
        return "Sustainability Index is Saved Successfully!", 201


class Sensitivity_Graph(Resource):

    def __init__(self):
        pass

    def get(self):
        """
        get the sensitivity index from files
        :return: sensitivity graph
        """
        node, link, index_value = get_data.get_data_coupled()
        # print(node, link)
        return {"node": node, "link": link, 'index-value': index_value}

class Load_Simulation_History(Resource):

    def __init__(self):
        pass

    def get(self):
        with open(".\\simulation_history\\simulation_history.json", "r") as file:
            simulation_history = json.load(file)
        return simulation_history, 200

    def post(self):
        simulation_history = request.get_json(self)
        with open(".\\simulation_history\\simulation_history.json", "w") as file:
            json.dump(simulation_history, file)
        return "Simulation result saved!", 201

# WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
# print(WEAP.ResultValue("Demand Sites and Catchments\\New_Magma\\Alfalfa_hay:Area Calculated", 2019, 1, "Reference",
# 					                  2019, 12, 'Average'))

