"""
For simulation results extraction
This module is the LEAP Visualization backend for results extraction
"""

import win32com.client
from mabia_model.script_m import MaxU
import pythoncom
import json
import pandas as pd
import numpy as np

def LEAP_visualization_variables():
	"""
	Run this function to geenrate all LEAP result variables' addresses and branch names
	:return:
	"""
	# "Unmet Requirements" "Efficiency", , "Self-sufficiency" and "Primary Requirements: Allocated to Demands" are not included since LEAP API is giving
	# error for those two variables.
	pythoncom.CoInitialize()
	Demand_Variables = ['Energy Demand Final Units', 'Load Shape']
	Transformation_Variables = ['Requirements', 'Outputs by Output Fuel', 'Outputs by Feedstock Fuel',
	                            'Inputs', 'Exports into Module',
	                            'Imports into Module', 'Capacity', 'Capacity Added',
	                            'Capacity Retired', 'Reserve Margin', 'Load Factor',
	                            'Peak Power Requirements', 'Actual Availability',
	                            'Power Generation', 'Energy Generation', 'Module Energy Balance']
	Resource_Variables = ['Reserves', 'Primary Requirements', 'Primary Supply', 'Indigenous Production', 'Imports', 'Exports']
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	# LEAP.ActiveArea = 'Ag_MABIA_v14'
	# active_scenario = ''

	# for s in LEAP.Scenarios:
	# 	print(s)
	# 	if s != 'Current Account':
	# 		active_scenario = s
	# LEAP.ActiveScenario = active_scenario
	start_year = LEAP.BaseYear
	LEAP_input = []
	LEAP_output = []
	# for v in LEAP.Branch('Demand\\Industrial\\Industrial water unrelated\\Industrial customers').Variables:
	# 	if v.name == 'Primary Requirements: Allocated to Demands':
	# 		print(v.name)
	# 		print(LEAP.Branch('Demand\\Industrial\\Industrial water unrelated\\Industrial customers').Variable(v.name).Value(2009))
	variables_to_visualize = {'Demand':{}, 'Transformation':{}, 'Resource': {}}
	for name in Demand_Variables:
		variables_to_visualize['Demand'][name] = []
	for name in Transformation_Variables:
		variables_to_visualize['Transformation'][name] = []
	for name in Resource_Variables:
		variables_to_visualize['Resource'][name] = []
	for b in LEAP.Branches:
		print('\n')
		print(b.FullName)
		print(b.Name)
		for v in LEAP.Branch(b.FullName).Variables:
			if v.name in Demand_Variables:
				try:
					LEAP.Branch(b.FullName).Variable(v.name).Value(start_year)
					variables_to_visualize['Demand'][v.name].append({'branch': b.FullName, 'variable': v.name})
				except:
					pass
			if v.name in Transformation_Variables:
				try:
					LEAP.Branch(b.FullName).Variable(v.name).Value(start_year)
					variables_to_visualize['Transformation'][v.name].append({'branch': b.FullName, 'variable': v.name})
				except:
					pass
			if v.name in Resource_Variables:
				try:
					LEAP.Branch(b.FullName).Variable(v.name).Value(start_year)
					variables_to_visualize['Resource'][v.name].append({'branch': b.FullName, 'variable': v.name})
				except:
					pass

	with open('D:\\Project\\Food_Energy_Water\\fewsim-backend\model\\LEAP_visualization_variables.json', 'w') as file:
		json.dump(variables_to_visualize, file)
	pythoncom.CoUninitialize()

def get_LEAP_value():
	"""
	This function extracts simulation results for all LEAP variables
	:return:
	"""
	pythoncom.CoInitialize()
	with open('D:\\Project\\Food_Energy_Water\\fewsim-backend\\model\\LEAP_visualization_variables.json', 'r') as file:
		variables = json.load(file)
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	# LEAP.ActiveArea = 'Ag_MABIA_v14'
	start_year = LEAP.BaseYear
	end_year = LEAP.EndYear
	for s in LEAP.Scenarios:
		if s != 'Current Account':
			active_scenario = s
	LEAP.ActiveScenario = active_scenario

	data = {}
	for i in variables.keys():
		data[i] = {}
		for j in variables[i].keys():
			data[i][j] = []
			for v in variables[i][j]:
				value_year = []
				for y in range(start_year + 1, end_year + 1):
					# print(LEAP.Branch(v['branch']).Variable(v['variable']).Value(y))
					print(v['branch'], v['variable'])
					value_year.append(LEAP.Branch(v['branch']).Variable(v['variable']).Value(y))
				data[i][j].append({'branch': v['branch'], 'variable': v['variable'], 'value': value_year})
	timeRange = [start_year + 1, end_year]
	# with open('.\model\LEAP_TEST_CACHE.json', 'r') as file:
	# 	data = json.load(file)
	print(data)
	pythoncom.CoUninitialize()
	return data, timeRange
# get_LEAP_value()
# LEAP_visualization_variables()
# LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
# LEAP.Branch("Transformation\Electricity generation").Variable("Energy Generation").Value(2020)