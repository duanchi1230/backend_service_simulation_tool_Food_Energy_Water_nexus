import win32com.client
import json
import logging
import pandas as pd
import numpy as np
from datetime import datetime
import pythoncom
from model import WEAP_Visualization_Model as WEAP_model
from model import LEAP_Visualization_Model as LEAP_model
from model import Food_Visualization_Model as food_model
import MPMmodel.transform_to_inputs as MPM_inputs
from climate_data.climate_data_fn import Climate_Data

def save_all_parameters():
	pythoncom.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	# start_year = WEAP.BaseYear
	# end_year = WEAP.EndYear
	WEAP.ActiveArea = 'Ag_MABIA_v14'
	WEAP.ActiveScenario = WEAP.Scenarios[1]
	weap_default_paramters = []
	mabia_catchment = pd.read_excel('Mabia_Catchments.xlsx', index_col=0)
	print(bool(np.intersect1d(['Tonopah'], np.array(mabia_catchment['variables']))))
	variable_list = pd.read_excel('WEAP_Input_Variables.xlsx')
	variable_list = np.array(variable_list['variable_name'])
	for b in WEAP.Branches:
		for v in WEAP.Branch(b.FullName).Variables:
			if v.name in variable_list:
				print(b.FullName, v.name, v.Expression)
				weap_default_paramters.append([b.FullName, v.name, v.Expression])
	weap_default_paramters = pd.DataFrame(weap_default_paramters, columns=['branch', 'variable', 'expression'])
	weap_default_paramters.to_csv('weap_default_paramters.csv')
	
	leap_default_paramters = []
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	input_variable_list = pd.read_excel('LEAP_Input_Variables.xlsx')
	input_variable_list = np.array(input_variable_list['variable_name'])
	for b in LEAP.Branches:
		for v in LEAP.Branch(b.FullName).Variables:
			if v.name in input_variable_list and b.FullName != 'Transformation\Electricity generation' and v.name != 'Unmet Requirements':
				print(b.FullName, v.name, v.Expression)
				leap_default_paramters.append([b.FullName, v.name, v.Expression])
	leap_default_paramters = pd.DataFrame(leap_default_paramters, columns=['branch', 'variable', 'expression'])
	leap_default_paramters.to_csv(('leap_default_paramters.csv'))
	print(leap_default_paramters)


def run_all_secanrios(scenarios, sustainability_variables, loaded_group_index):
	pythoncom.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	# WEAP.ActiveArea = 'Ag_MABIA_v14'
	WEAP.ActiveScenario = list(WEAP.Scenarios)[1]
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	# LEAP.ActiveArea = 'Ag_MABIA_v14'
	LEAP.ActiveScenario = list(LEAP.Scenarios)[1]
	default_scenarios = []
	weap_flow = []
	leap_data = []
	food_data = []
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	MPM = MPM_inputs.MPM()
	MPM.set_MPM_MABIA(WEAP)
	################################################################################################################
	now = datetime.now()
	dt_string = now.strftime('%d/%m/%Y %H:%M:%S')
	message = 'Simulation is initializaing!'
	log_file = pd.DataFrame([[dt_string, message]], columns=['time', 'message'])
	log_file.to_csv('log.csv')
	# run_log_file = open('run_log_file.txt', 'w')
	################################################################################################################
	# WEAP.Calculate()
	# log_file = append_log(log_file, 'WEAP completed running Scenario Base')
	# ################################################################################################################
	# log_file = append_log(log_file, 'LEAP started running Scenario Base')
	# LEAP.Calculate()
	# log_file = append_log(log_file, 'LEAP completed running Scenario Base')
	# ################################################################################################################
	# log_file = append_log(log_file, 'WEAP started extracting Scenario Base results')
	# flow, timeRange_WEAP = WEAP_model.get_WEAP_flow_value()
	# value = flow[list(flow.keys())[0]]
	# weap_flow.append({
	# 	'sid': 0,
	# 	'name': 'Base',
	# 	'runStatus': 'finished',
	# 	'timeRange': timeRange_WEAP,
	# 	'numTimeSteps': timeRange_WEAP[1] - timeRange_WEAP[0],
	# 	'var': {'input': 'base_parameters',
	# 	        'output': value}
	# })
	# log_file = append_log(log_file, 'WEAP completed extracting Scenario Base results')
	# ################################################################################################################
	# log_file = append_log(log_file, 'LEAP started extracting Scenario Base results')
	# data, timeRange_LEAP = LEAP_model.get_LEAP_value()
	# leap_data.append({
	# 	'sid': 0,
	# 	'name': 'Base',
	# 	'runStatus': 'finished',
	# 	'timeRange': timeRange_LEAP,
	# 	'numTimeSteps': timeRange_LEAP[1] - timeRange_LEAP[0],
	# 	'var': {'input': 'base_parameters',
	# 	        'output': data}
	# })
	# log_file = append_log(log_file, 'LEAP completed extracting Scenario Base results')
	# ################################################################################################################
	# print('created scenarios--->', scenarios)
	# ################################################################################################################
	##################################################################################################################
	# process sustainability variables
	for i in range(len(sustainability_variables)):
		sustainability_variables[i]["node"]["value"] = []
	
	for i in range(len(loaded_group_index)):
		for j in range(len(loaded_group_index[i]["variable"])):
			loaded_group_index[i]["variable"][j]["node"]["value"] = []
	###################################################################################################################
	for scenario in scenarios:
		WEAP.View = 'Data'
		LEAP.ActiveView = 'Analysis'
		default_variable_weap = []
		default_variable_leap = []
		default_variable_mabia = []
		
		for variable in scenario['weap']:
			expression = str(WEAP.Branch(variable['fullname']).Variable(variable['name']).Expression) + '*' + str(
				variable['percentage_of_default']) + '%'
			default = WEAP.Branch(variable['fullname']).Variable(variable['name']).Expression
			default_variable_weap.append({'branch': variable['fullname'], 'name': variable['name'],
			                              'expression': default,
			                              'percentage_of_default': variable['percentage_of_default'], 'model': 'weap'})
			for s in WEAP.Scenarios:
				WEAP.ActiveScenario = s
				WEAP.Branch(variable['fullname']).Variable(variable['name']).Expression = expression
			print('setting weap variable--->', variable['fullname'], '--->', variable['name'], '--->', expression)
			print('saving weap default--->', variable['fullname'], '--->', variable['name'], '--->', default)
		for variable in scenario['leap']:
			expression = str(LEAP.Branch(variable['fullname']).Variable(variable['name']).Expression) + '*' + str(
				variable['percentage_of_default']) + '%'
			default = LEAP.Branch(variable['fullname']).Variable(variable['name']).Expression
			default_variable_leap.append({'branch': variable['fullname'], 'name': variable['name'],
			                              'expression': default,
			                              'percentage_of_default': variable['percentage_of_default'], 'model': 'leap'})
			for s in LEAP.Scenarios:
				LEAP.ActiveScenario = s
				LEAP.Branch(variable['fullname']).Variable(variable['name']).Expression = expression
			print('setting leap variable--->', variable['fullname'], '--->', variable['name'], '--->', expression)
			print('saving leap default--->', variable['fullname'], '--->', variable['name'], '--->', default)
		for variable in scenario['climate']:
			climate_scenario_name = variable.split("_")[0]
			climate_scenatio_type = variable.split("_")[1]
			cimate_file_handler = Climate_Data()
			cimate_file_handler.set_climate_MABIA(scenario_name=climate_scenario_name, scenatio_type=climate_scenatio_type)

			# handle the mpm files associated with climate scenario
			climate_path = pd.read_csv("D:\\Project\\Food_Energy_Water\\fewsim-backend\\MPMmodel\\climate_file_macth.csv", index_col=0)
			if climate_scenatio_type != "Hist":
				climate_input = climate_path.loc[(climate_path["type"] == climate_scenatio_type) & (climate_path["climate_scenario"] == climate_scenario_name)]["mpm_path"].iloc[0]
				MPM.set_MPM_climate(WEAP, climate_input=climate_input)
			if climate_scenatio_type != "Hist":
				MPM.set_MPM_default(WEAP)
		# for variable in scenario['mabia']:
		#
		# 	default_variable_mabia.append({'branch': variable['branch'], 'name': variable['name'],
		# 	                               'expression': LEAP.Branch(variable['branch']).Variable(
		# 		                               variable['name']).Expression,
		# 	                               'percentage_of_default': variable['percentage_of_default'],
		# 	                               'model': 'mabia'})
		#
		# 	WEAP.Branch(variable['branch']).Variable(variable['name']).Expression = variable['expression']


		################################################################################################################
		log_file = append_log(log_file, 'WEAP started running Scenario ' + scenario['name'])
		WEAP.Calculate()
		log_file = append_log(log_file, 'WEAP completed running Scenario ' + scenario['name'])
		################################################################################################################
		log_file = append_log(log_file, 'LEAP started running Scenario ' + scenario['name'])
		LEAP.Calculate()
		log_file = append_log(log_file, 'LEAP completed running Scenario ' + scenario['name'])
		################################################################################################################
		log_file = append_log(log_file, 'WEAP started extracting Scenario results from ' + scenario['name'])
		flow, timeRange_WEAP = WEAP_model.get_WEAP_flow_value()
		weap_result = flow[list(flow.keys())[0]]
		weap_flow.append({
			'sid': 0,
			'name': scenario['name'],
			'runStatus': 'finished',
			'timeRange': timeRange_WEAP,
			'numTimeSteps': timeRange_WEAP[1] - timeRange_WEAP[0],
			'var': {'input': scenario,
			        'output': weap_result}
		})
		log_file = append_log(log_file, 'WEAP completed extracting Scenario results from ' + scenario['name'])
		################################################################################################################
		log_file = append_log(log_file, 'LEAP started extracting Scenario results from ' + scenario['name'])
		leap_result, timeRange_LEAP = LEAP_model.get_LEAP_value()
		leap_data.append({
			'sid': 0,
			'name': scenario['name'],
			'runStatus': 'finished',
			'timeRange': timeRange_LEAP,
			'numTimeSteps': timeRange_LEAP[1] - timeRange_LEAP[0],
			'var': {'input': scenario,
			        'output': leap_result}
		})
		log_file = append_log(log_file, 'LEAP completed extracting Scenario results from ' + scenario['name'])
		print('run_scenarios: 150')
		################################################################################################################
		log_file = append_log(log_file, 'Extracting food variables in  ' + scenario['name'])
		food_result = food_model.get_food_variables()
		food_data.append({
			'sid': 0,
			'name': scenario['name'],
			'runStatus': 'finished',
			'timeRange': timeRange_LEAP,
			'numTimeSteps': timeRange_LEAP[1] - timeRange_LEAP[0],
			'var': {'input': scenario,
			        'output': food_result}
		})
		log_file = append_log(log_file, 'Completed extracting food variables from Scenario ' + scenario['name'])
		################################################################################################################
		# extract (customized) sustainability variables
		LEAP.ActiveView = 'Results'
		WEAP.View = 'Results'
		for i in range(len(sustainability_variables)):
			if sustainability_variables[i]["node"]["model"] == "weap":
				s_weap_variable = []
				for year in range(timeRange_WEAP[0], timeRange_WEAP[1] + 1):
					s_weap_variable.append(WEAP.ResultValue(sustainability_variables[i]["variable"], year, 1, "Reference",
						                            year, 12, 'Total'))
				sustainability_variables[i]["node"]["value"].append({"name": scenario["name"], "calculated":s_weap_variable} )
			if sustainability_variables[i]["node"]["model"] == "leap":
				for s in LEAP.Scenarios:
					if s != 'Current Account':
						active_scenario = s
				LEAP.ActiveScenario = active_scenario
				s_leap_variable = []
				print(sustainability_variables[i]["node"]["name"])
				for year in range(timeRange_LEAP[0], timeRange_LEAP[1] + 1):
					s_leap_variable.append(LEAP.Branch(sustainability_variables[i]["node"]["fullname"]).Variable(sustainability_variables[i]["node"]["name"]).Value(year))
				sustainability_variables[i]["node"]["value"].append({"name": scenario["name"], "calculated":s_leap_variable})
			if sustainability_variables[i]["node"]["model"] == "mabia":
				s_mabia_variable = []
				for year in range(timeRange_WEAP[0], timeRange_WEAP[1] + 1):
					s_mabia_variable.append(WEAP.ResultValue(sustainability_variables[i]["variable"], year, 1, "Reference", year, 12, 'Total'))
				sustainability_variables[i]["node"]["value"].append(
					{"name": scenario["name"], "calculated":s_mabia_variable})

			if sustainability_variables[i]["node"]["model"] == "mpm":
				total_Croprea = 439100836.4
				root_path = "D:\Project\Food_Energy_Water\\fewsim-backend"
				mpm_outputs = pd.read_csv(root_path + "\MPMmodel\outPuts.csv", index_col=0)
				crops = {'cotton': "0", 'alfalfa': "1", 'corn': "2", 'barley': "3", 'durum': "4", 'veg': "5", 'remaining': "6"}
				col = crops[sustainability_variables[i]["node"]["name"]]
				mpm_result = (
						mpm_outputs.loc[start_year + 1:end_year, col].to_numpy() * total_Croprea).tolist()
				sustainability_variables[i]["node"]["value"].append(
					{"name": scenario["name"], "calculated": mpm_result})
		################################################################################################################
		# extract loaded (saved) sustainability variables
		LEAP.ActiveView = 'Results'
		WEAP.View = 'Results'
		for i in range(len(loaded_group_index)):
			for j in range(len(loaded_group_index[i]["variable"])):
				if loaded_group_index[i]["variable"][j]["node"]["model"] == "weap":
					s_weap_variable = []
					for year in range(timeRange_WEAP[0], timeRange_WEAP[1] + 1):
						s_weap_variable.append(
							WEAP.ResultValue(loaded_group_index[i]["variable"][j]["variable"], year, 1, "Reference",
							                 year, 12, 'Total'))
					loaded_group_index[i]["variable"][j]["node"]["value"].append(
						{"name": scenario["name"], "calculated": s_weap_variable})
				if loaded_group_index[i]["variable"][j]["node"]["model"] == "leap":
					for s in LEAP.Scenarios:
						if s != 'Current Account':
							active_scenario = s
					LEAP.ActiveScenario = active_scenario
					s_leap_variable = []
					# print(sustainability_variables[i]["node"]["name"])
					for year in range(timeRange_LEAP[0], timeRange_LEAP[1] + 1):
						s_leap_variable.append(
							LEAP.Branch(loaded_group_index[i]["variable"][j]["node"]["fullname"]).Variable(
								loaded_group_index[i]["variable"][j]["node"]["name"]).Value(year))
					loaded_group_index[i]["variable"][j]["node"]["value"].append(
						{"name": scenario["name"], "calculated": s_leap_variable})
				if loaded_group_index[i]["variable"][j]["node"]["model"] == "mabia":
					s_mabia_variable = []
					for year in range(timeRange_WEAP[0], timeRange_WEAP[1] + 1):
						s_mabia_variable.append(
							WEAP.ResultValue(loaded_group_index[i]["variable"][j]["variable"], year, 1, "Reference", year,
							                 12, 'Total'))
					loaded_group_index[i]["variable"][j]["node"]["value"].append(
						{"name": scenario["name"], "calculated": s_mabia_variable})

				if loaded_group_index[i]["variable"][j]["node"]["model"] == "mpm":
					total_Croprea = 439100836.4
					root_path = "D:\Project\Food_Energy_Water\\fewsim-backend"
					mpm_outputs = pd.read_csv(root_path + "\MPMmodel\outPuts.csv", index_col=0)
					crops = {'cotton':"0", 'alfalfa':"1", 'corn':"2", 'barley':"3", 'durum':"4", 'veg':"5", 'remaining':"6"}
					col = crops[loaded_group_index[i]["variable"][j]["node"]["name"]]
					mpm_result = (
								mpm_outputs.loc[start_year + 1:end_year, col].to_numpy() * total_Croprea).tolist()
					loaded_group_index[i]["variable"][j]["node"]["value"].append(
						{"name": scenario["name"], "calculated": mpm_result})
		################################################################################################################
		# set variables back to default
		# set WEAP back to default
		WEAP.View = 'Data'
		for default_variable in default_variable_weap:
			for s in WEAP.Scenarios:
				WEAP.ActiveScenario = s
				WEAP.Branch(default_variable['branch']).Variable(default_variable['name']).Expression = \
				default_variable['expression']
			print('setting weap default--->', default_variable['branch'], '--->', default_variable['name'], '--->',
			      default_variable['expression'])
		# set LEAP back to default
		LEAP.ActiveView = 'Analysis'
		for default_variable_l in default_variable_leap:
			for s in LEAP.Scenarios:
				LEAP.ActiveScenario = s
				LEAP.Branch(default_variable_l['branch']).Variable(default_variable_l['name']).Expression = \
				default_variable_l['expression']
			print('setting leap default--->', default_variable_l['branch'], '--->', default_variable_l['name'], '--->',
			      default_variable_l['expression'])
		# set climate files and mpm files back to default
		if len(scenario['climate'])>0:
			cimate_file_handler.set_climate_default()
			MPM.set_MPM_default(WEAP)
	log_file = append_log(log_file, 'Completed')
	###################################################################################################################
	# calculate weap_flow response level in percentage
	max_w = {}
	w = {}
	for f in weap_flow:
		for v in f["var"]["output"]:
			print(v)
			w[v["name"]] = []
		for v in f["var"]["output"]:
			w[v["name"]].append(v["value"])
			max_w[v["name"]] = np.amax(w[v["name"]], axis=0)
	print(max_w)
	for i in range(len(weap_flow)):
		for j in range(len(weap_flow[i]["var"]["output"])):
			weap_flow[i]["var"]["output"][j]["percentage"] = list(np.around((np.array(
				weap_flow[i]["var"]["output"][j]["value"]) - np.array(weap_flow[0]["var"]["output"][j]["value"])) / (
						                                                                max_w[weap_flow[i]["var"][
							                                                                "output"][j][
							                                                                "name"]] + 0.000001),
			                                                                decimals=3))
	print(weap_flow)
	###################################################################################################################
	# calculate leap_data response level in percentage
	max_l = {}
	l = {}
	for type in leap_data[0]["var"]["output"].keys():
		l[type] = {}
		max_l[type] = {}
		for variable in leap_data[0]["var"]["output"][type].keys():
			l[type][variable] = {}
			max_l[type][variable] = {}
			for branch in leap_data[0]["var"]["output"][type][variable]:
				l[type][variable][branch["branch"]] = []
	for d in leap_data:
		for type in d["var"]["output"].keys():
			for variable in d["var"]["output"][type].keys():
				for branch in d["var"]["output"][type][variable]:
					l[type][variable][branch["branch"]].append(branch["value"])
					max_l[type][variable][branch["branch"]] = np.amax(l[type][variable][branch["branch"]], axis=0)
	for i in range(len(leap_data)):
		for type in leap_data[i]["var"]["output"].keys():
			for variable in leap_data[i]["var"]["output"][type].keys():
				for j in range(len(leap_data[i]["var"]["output"][type][variable])):
					l_i = leap_data[i]["var"]["output"][type][variable][j]["value"]
					l_0 = leap_data[0]["var"]["output"][type][variable][j]["value"]
					max_l_i = max_l[type][variable][leap_data[i]["var"]["output"][type][variable][j]["branch"]]
					leap_data[i]["var"]["output"][type][variable][j]["percentage"] = list(
						np.around((np.array(l_i) - np.array(l_0)) / (np.array(max_l_i) + 0.000001), decimals=3))

	pythoncom.CoUninitialize()

	with open('run_results.json', 'w') as outfile:
		json.dump({'weap-flow': weap_flow, 'leap-data': leap_data,  "food-data":food_data, "sustainability_variables": sustainability_variables, "loaded_group_index": loaded_group_index}, outfile)
	return weap_flow, leap_data, food_data, sustainability_variables, loaded_group_index


def append_log(log_file, message):
	now = datetime.now()
	dt_string = now.strftime('%d/%m/%Y %H:%M:%S')
	log_file = pd.concat([log_file, pd.DataFrame([[dt_string, message]], columns=['time', 'message'])],
	                     ignore_index=True)
	log_file.to_csv('log.csv')
	return log_file

# scenarios = []
# weap_flow, leap_data = run_all_secanrios(scenarios)
# print(weap_flow,leap_data)
#
# log_file = pd.DataFrame([[0, 'asdasd']],columns=['time', 'message'])
# log_file = pd.concat([log_file, pd.DataFrame([[1, 'qweqe']],columns=['time', 'message'])], ignore_index=True)
# for row in log_file.iterrows():
# 	print(row[1]['time'], row[1]['message'])
# climate_path = pd.read_csv("D:\\Project\\Food_Energy_Water\\fewsim-backend\\MPMmodel\\climate_file_macth.csv", index_col=0)
# print(climate_path.loc[(climate_path["type"] == "ssp585") & (climate_path["climate_scenario"] == "CanESM5")]["mpm_path"].iloc[0])