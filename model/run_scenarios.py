import win32com.client
import json
import logging
import pandas as pd
import numpy as np
from datetime import datetime
import pythoncom
from model import WEAP_Visualization_Model as WEAP_model
from model import LEAP_Visualization_Model as LEAP_model


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


def run_all_secanrios(scenarios):
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
	################################################################################################################
	now = datetime.now()
	dt_string = now.strftime('%d/%m/%Y %H:%M:%S')
	message = 'Simulation is initializaing!'
	log_file = pd.DataFrame([[dt_string, message]],columns=['time', 'message'])
	log_file.to_csv('log.csv')
	run_log_file = open('run_log_file.txt', 'w')
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
					variable['percentage_of_default'])+'%'
			default = LEAP.Branch(variable['fullname']).Variable(variable['name']).Expression
			default_variable_leap.append({'branch': variable['fullname'], 'name': variable['name'],
			                              'expression': default,
			                              'percentage_of_default': variable['percentage_of_default'], 'model': 'leap'})
			for s in LEAP.Scenarios:
				LEAP.ActiveScenario = s
				LEAP.Branch(variable['fullname']).Variable(variable['name']).Expression = expression
			print('setting leap variable--->', variable['fullname'], '--->', variable['name'], '--->', expression)
			print('saving leap default--->', variable['fullname'], '--->', variable['name'], '--->', default)
			
		# for variable in scenario['mabia']:
		#
		# 	default_variable_mabia.append({'branch': variable['branch'], 'name': variable['name'],
		# 	                               'expression': LEAP.Branch(variable['branch']).Variable(
		# 		                               variable['name']).Expression,
		# 	                               'percentage_of_default': variable['percentage_of_default'],
		# 	                               'model': 'mabia'})
		#
		# 	LEAP.Branch(variable['branch']).Variable(variable['name']).Expression = variable['expression']
		
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
		log_file = append_log(log_file, 'LEAP completed extracting Scenario results from ' + scenario['name'])
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
		WEAP.View = 'Data'
		for default_variable in default_variable_weap:
			for s in WEAP.Scenarios:
				WEAP.ActiveScenario = s
				WEAP.Branch(default_variable['branch']).Variable(default_variable['name']).Expression = default_variable['expression']
			print('setting weap default--->', default_variable['branch'], '--->', default_variable['name'], '--->',
		      default_variable['expression'])
		LEAP.ActiveView = 'Analysis'
		for default_variable_l in default_variable_leap:
			for s in LEAP.Scenarios:
				LEAP.ActiveScenario = s
				LEAP.Branch(default_variable_l['branch']).Variable(default_variable_l['name']).Expression = default_variable_l['expression']
			print('setting leap default--->', default_variable_l['branch'], '--->', default_variable_l['name'], '--->',
		      default_variable_l['expression'])
	log_file = append_log(log_file, 'Completed')
	pythoncom.CoUninitialize()
	with open('run_results.json', 'w') as outfile:
		json.dump({'weap-flow': weap_flow, 'leap-data': leap_data}, outfile)
	return weap_flow, leap_data

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