import win32com.client
import json
import pandas as pd
import numpy as np
import time
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
	WEAP.ActiveArea = 'Ag_MABIA_v14'
	WEAP.ActiveScenario = list(WEAP.Scenarios)[1]
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	LEAP.ActiveArea = 'Ag_MABIA_v14'
	LEAP.ActiveScenario = list(LEAP.Scenarios)[1]
	default_scenarios = []
	weap_flow =[]
	leap_data = []
	# WEAP.Calculate()
	flow, timeRange_WEAP = WEAP_model.get_WEAP_flow_value()
	value = flow[list(flow.keys())[0]]
	weap_flow.append({
		'sid': 0,
		'name': 'Base',
		'runStatus': 'finished',
		'timeRange': timeRange_WEAP,
		'numTimeSteps': timeRange_WEAP[1] - timeRange_WEAP[0],
		'var': {'input': 'base_parameters',
		        'output': value}
	})

	# LEAP.Calculate()
	data, timeRange_LEAP = LEAP_model.get_LEAP_value()
	leap_data.append({
		'sid': 0,
		'name': 'Base',
		'runStatus': 'finished',
		'timeRange': timeRange_LEAP,
		'numTimeSteps': timeRange_LEAP[1] - timeRange_LEAP[0],
		'var': {'input': 'base_parameters',
		        'output': data}
	})

	for scenario in scenarios:
		default_variable = []
		for variable in scenario['policy']:
			if variable['model'] == 'weap':
				default_variable.append({'branch': variable['branch'], 'name': variable['name'],
				                         'expression': WEAP.Branch(variable['branch']).Variable(
					                         variable['name']).Expression, 'model': 'weap'})
				WEAP.Branch(variable['branch']).Variable(variable['name']).Expression = variable['expression']
			if variable['model'] == 'mabia':
				default_variable.append({'branch': variable['branch'], 'name': variable['name'],
				                         'expression': WEAP.Branch(variable['branch']).Variable(
					                         variable['name']).Expression, 'model': 'weap'})
				WEAP.Branch(variable['branch']).Variable(variable['name']).Expression = variable['expression']
			if variable['model'] == 'leap':
				default_variable.append({'branch': variable['branch'], 'name': variable['name'],
				                         'expression': LEAP.Branch(variable['branch']).Variable(
					                         variable['name']).Expression, 'model': 'weap'})
				LEAP.Branch(variable['branch']).Variable(variable['name']).Expression = variable['expression']
		default_scenarios.append(default_variable)
		WEAP.Calculate()
		flow, timeRange_WEAP = WEAP_model.get_WEAP_flow_value()
		value = flow[list(flow.keys())[0]]
		weap_flow.append({
			'sid': 0,
			'name': scenario['name'],
			'runStatus': 'finished',
			'timeRange': timeRange_WEAP,
			'numTimeSteps': timeRange_WEAP[1] - timeRange_WEAP[0],
			'var': {'input': scenario,
			        'output': value}
		})

		data, timeRange_LEAP = LEAP_model.get_LEAP_value()
		leap_data.append({
			'sid': 0,
			'name': scenario['name'],
			'runStatus': 'finished',
			'timeRange': timeRange_LEAP,
			'numTimeSteps': timeRange_LEAP[1] - timeRange_LEAP[0],
			'var': {'input': scenario,
			        'output': value}
		})

		# LEAP.Calculate()

	pythoncom.CoUninitialize()
	return weap_flow, leap_data

# scenarios = []
# weap_flow, leap_data = run_all_secanrios(scenarios)
# print(weap_flow,leap_data)
