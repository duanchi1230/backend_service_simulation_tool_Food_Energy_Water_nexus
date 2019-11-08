import win32com.client
import json
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
"""
	THIS MODULE IS THE WEAP-BACKEND FOR FEWSIM SYSTEM
"""

def get_WEAP_variables():
	"""
	This function extract all results values from WEAP
	:return: Structured dictionary of WEAP results value
	"""
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	WEAP.ActiveArea = 'Internal_Linking_test_das'
	WEAP.ActiveScenario = WEAP.Scenarios[1]
	WEAP_input = []
	WEAP_output = []
	for b in WEAP.Branches:
		WEAP.Branch(b.FullName)
		# print('\n')
		# print(b.FullName)
		# print(b.Name)
		for v in WEAP.Branch(b.FullName).Variables:
			if v.isResultVariable:
				value = []
				if v != None:
					for y in range(start_year, end_year + 1):
						path = b.FullName + ":" + v.name
						print(WEAP.ResultValue(path, start_year, 1, 'Linkage', end_year, 12, 'Total'))
						value.append(WEAP.ResultValue(path, start_year, 1, 'Linkage', end_year, 12, 'Total'))
					unit = WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit
					path = path_parser(b.FullName)
					path.append(v.name)
					node = {
						'name': v.name,
						'fullname': b.FullName,
						'path': path,
						'parent': path[-2] if len(path) > 1 else 'null',
						'value': value,
						'unit': unit
					}
					WEAP_output = tree_insert_node(path, node, WEAP_output)

			variable_list = pd.read_excel('WEAP_Input_Variables.xlsx')
			variable_list = np.array(variable_list['variable_name'])
			if v.name in variable_list:
				print(b.FullName, v.name)
				value = []
				for y in range(start_year, end_year + 1):
					path = b.FullName + ":" + v.name
					# print(WEAP.ResultValue(path, y, 1, 'Linkage', y, 12, 'Average'), WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit)
					value.append(WEAP.ResultValue(path, y, 1, 'Linkage', y, 12, 'Average'))
				unit = WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit
				path = path_parser(b.FullName)
				path.append(v.name)
				node = {
					'name': v.name,
					'fullname': b.FullName,
					'path': path,
					'parent': path[-2] if len(path) > 1 else 'null',
					'value': value,
					'unit': unit
				}
				WEAP_input = tree_insert_node(path, node, WEAP_input)
				print(WEAP_input)
	with open('WEAP_variables.json', 'w') as outfile:
		json.dump({'WEAP-input':WEAP_input, 'WEAP-output': WEAP_output}, outfile)

	win32com.CoUninitialize()

def path_parser(path):
	"""
	:param path: A string in the format of 'Transformation\Electricity generation\Output Fuels\Electricity'
	:return: A parsed string array in the format ['Transformation', 'Electricity generation', 'Output Fuels', 'Electricity']
	"""
	branch = []
	name = ''
	for character in path:
		if character != '\\':
			name = name + character
		else:
			if name != '':
				branch.append(name)
				name = ''
	branch.append(name)
	return branch


path = 'Transformation\Electricity generation\Output Fuels\Electricity'
# print(path_parser(path))

def tree_find_key(path_key, tree):
	path = 'tree'
	for key in path_key:
		i = 0
		while i<len(eval(path)):
			try:
				if eval(path)[i]['name'] == key:
					if key != path_key[-1]:
						path = path + '[' + str(i) +']' + "['children']"
					else:
						path = path + '[' + str(i) + ']'
			except:
				pass
			i = i + 1
		# print(eval(path))
	return eval(path)

def tree_insert_node(path_key, node, tree):
	path = 'tree'
	for key in path_key:
		i = 0
		key_exist = False
		# If the key exists in the current level, the search will update the path
		while i<len(eval(path)):
			try:
				if eval(path)[i]['name'] == key:
					key_exist = True
					if key != path_key[-1]:
						path = path + '[' + str(i) + ']' + "['children']"
					else:
						# path = path + '[' + str(i) + ']' + "['children']" + "=node"
						# eval(path)
						pass
			except:
				pass
			i = i + 1
		# If the key doesn't exist in current level, a new node will be inserted
		if key_exist == False:
			if key != path_key[-1]:
				intermediate_noden = {"name": key,
				     "parent": path_key[path_key.index(key) - 1] if key!=path_key[0] else 'null',
				     "children": []}
				eval(path).append(intermediate_noden)
				path = path + '[' + str(len(eval(path))-1) + ']' + "['children']"
			else:
				eval(path).append(node)
		# print(path)
	return tree

def get_WEAP_inputs():

	pass
get_WEAP_variables()


# WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')

# print(WEAP.Branch('Demand Sites and Catchments\\Agricultural Catchment\\winter_wheat').Variables('Area Calculated').Value)
# print(WEAP.ResultValue('\Demand Sites and Catchments\Municipal: Annual Activity Level', 2002, 1, 'Linkage', 2002,12, 'Total'))
# for v in WEAP.Branch('\Demand Sites and Catchments\Municipal').Variables:
# 	print(v.name, v.isResultVariable)
# for v in WEAP.Branch('Demand Sites and Catchments\Mun
# icipal').Variables:
# 	if v.IsResultVariable == True:
# 		print(v.Name, v.Value)