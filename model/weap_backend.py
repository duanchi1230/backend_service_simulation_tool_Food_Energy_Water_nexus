"""
This script can generate all variable addresses and branch names in the WEAP model
"""

import win32com.client
import json
import pandas as pd
import numpy as np
import time
import pythoncom
from pandas import ExcelWriter
from pandas import ExcelFile


def generate_WEAP_variables():
	"""
	This function extract all results values from WEAP
	:return: Structured dictionary of WEAP results value
	"""

	pythoncom.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	# WEAP.ActiveArea = 'Ag_MABIA_v14'
	WEAP.ActiveScenario = WEAP.Scenarios[1]
	WEAP_input = []
	WEAP_output = []
	mabia_catchment = pd.read_excel('Mabia_Catchments.xlsx', index_col=0)
	print(bool(np.intersect1d(["Tonopah"], np.array(mabia_catchment['variables']))))
	for b in WEAP.Branches:
		WEAP.Branch(b.FullName)
		# print('\n')
		print(b.FullName)
		# print(b.Name)
		for v in WEAP.Branch(b.FullName).Variables:
			if v.isResultVariable:
				value = []
				year = []
				if v != None:
					for y in range(start_year, end_year + 1):
						path = b.FullName + ":" + v.name
						print(WEAP.ResultValue(path, start_year, 1, WEAP.Scenarios[1], end_year, 12, 'Total'))
						value.append(WEAP.ResultValue(path, start_year, 1, WEAP.Scenarios[1], end_year, 12, 'Total'))
						year.append(y)
					unit = WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit
					path = path_parser(b.FullName)
					path.append(v.name)
					node = {
						'name': v.name,
						'fullname': b.FullName,
						'path': path,
						'parent': path[-2] if len(path) > 1 else 'null',
						'unit': unit,
						'year': year,
						'value': value
					}
					if len(np.intersect1d(path, np.array(mabia_catchment['variables'])))>0:
						node['model'] = 'mabia'
					else:
						node['model'] = 'weap'
					type_of_variable = 'output'
					WEAP_output = tree_insert_node(path, node, type_of_variable, WEAP_output)

			variable_list = pd.read_excel('WEAP_Input_Variables.xlsx')
			variable_list = np.array(variable_list['variable_name'])
			if v.name in variable_list:
				print(b.FullName, v.name)
				value = []
				for y in range(start_year, end_year + 1):
					path = b.FullName + ":" + v.name
					# print(WEAP.ResultValue(path, y, 1, 'Linkage', y, 12, 'Average'), WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit)
					value.append(WEAP.ResultValue(path, y, 1, WEAP.Scenarios[1], y, 12, 'Average'))
				unit = WEAP.Branch(b.FullName).Variable(v.name).ScaleUnit
				path = path_parser(b.FullName)
				path.append(v.name)
				node = {
					'name': v.name,
					'fullname': b.FullName,
					'path': path,
					'parent': path[-2] if len(path) > 1 else 'null',
					'model': 'weap',
					'value': value,
					'unit': unit
				}
				if len(np.intersect1d(path, np.array(mabia_catchment['variables'])))>0:
					node['model'] = 'mabia'
				else:
					node['model'] = 'weap'
				type_of_variable = 'input',
				WEAP_input = tree_insert_node(path, node, type_of_variable, WEAP_input)
				print(WEAP_input)
	with open('WEAP_variables.json', 'w') as outfile:
		json.dump([{'name': 'weap-input', 'children': WEAP_input},
		           {'name': 'weap-output', 'children': WEAP_output}] , outfile)

	pythoncom.CoUninitialize()


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
	"""
		This module is used to query the variable tree.
		:param path_key: The path for the node to be queried.
		:param tree: The tree to be query from.
		:return: The node value.
	"""
	path = 'tree'
	for key in path_key:
		i = 0
		while i < len(eval(path)):
			try:
				if eval(path)[i]['name'] == key:
					if key != path_key[-1]:
						path = path + '[' + str(i) + ']' + "['children']"
					else:
						path = path + '[' + str(i) + ']'
			except:
				pass
			i = i + 1
	# print(eval(path))
	return eval(path)


def tree_insert_node(path_key, node, type_of_variable, tree):
	"""
		This module is used to insert a node to the tree.
		:param path_key: The path of a node to be inserted.
		:param node: The node to be inserted.
		:param tree: The tree to which the node is inserted.
		:return: The tree with new nodes inserted.
	"""
	path = 'tree'
	mabia_catchment = pd.read_excel('Mabia_Catchments.xlsx', index_col=0)
	if len(np.intersect1d([node['name']], np.array(mabia_catchment['variables'])))>0:
		model = 'mabia'
	else:
		model = 'weap'
	for key in path_key:
		i = 0
		key_exist = False
		# If the key exists in the current level, the search will update the path
		while i < len(eval(path)):
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
				                      "parent": path_key[path_key.index(key) - 1] if key != path_key[0] else 'null',
				                      "model": model,
				                      "type": type_of_variable,
				                      "children": []}
				eval(path).append(intermediate_noden)
				path = path + '[' + str(len(eval(path)) - 1) + ']' + "['children']"
			else:
				eval(path).append(node)
	# print(path)
	return tree


def expand_tree(tree, input_list):
	"""
	This module decomposes a tree and find all the leaves of the tree.
	This method is using depth-first search.
	:param tree:  A tree with nodes and leaves
	:param input_list: an empty list that is used to hold the tree leaves
	:return: A list of leaves of the tree
	"""
	for v in tree:
		if 'children' not in v.keys():
			# print(v['fullname'], v['name'], v['value'])
			input_list.append(v)
		else:
			expand_tree(v['children'], input_list)
	return input_list


def get_WEAP_variables_from_file(file_path):
	"""
	This module grabs the list of WEAP variables and their paths from the stored local JSON file
	:param file_path: The path of the local file
	:return: input_list of all the WEAP inputs
	"""
	with open(file_path) as f:
		variables = json.load(f)
	variables_list = []
	print(variables[0].keys())
	variables_list = expand_tree(variables, variables_list)
	print(variables_list[0])
	return variables_list

def get_WEAP_variables_tree(file_path):
	with open(file_path) as f:
		variables = json.load(f)
	return variables

def list_variables(variables):
	input_list = pd.read_excel('WEAP_Input_Variables.xlsx')
	input_list = np.array(input_list['variable_name'])
	list = []
	for v in variables:
		if v['name'] in input_list:
			list.append([v['fullname'], v['name'], 'input'])
		else:
			list.append([v['fullname'], v['name'], 'output'])
	df = pd.DataFrame(list, columns=['branch', 'variable-name', 'type'])
	df.to_csv('W_variables.csv')
	print(df)

# get_WEAP_inputs_tree('WEAP_variables.json')
# variables = get_WEAP_variables_from_file('WEAP_variables.json')
# list_variables(variables)

# start_time = time.time()
# generate_WEAP_variables()
# elapsed_time = time.time() - start_time
# print('Extraction of all WEAP variables takes: ',elapsed_time, ' s')

# WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
# print(WEAP.Branch('Demand Sites and Catchments\\Agricultural Catchment\\winter_wheat').Variables('Area Calculated').Value)
# print(WEAP.ResultValue('\Demand Sites and Catchments\Municipal: Water Demand', 2002, 1, 'Linkage', 2002,12, 'Total'))
# for v in WEAP.Branch('\Demand Sites and Catchments\Municipal').Variables:
# 	print(v.name)
# for v in WEAP.Branch('\Demand Sites and Catchments\Municipal').Variables:
# 	print(v.name, v.isResultVariable)
# for v in WEAP.Branch('Demand Sites and Catchments\Mun
# icipal').Variables:
# 	if v.IsResultVariable == True:
# 		print(v.Name, v.Value)
