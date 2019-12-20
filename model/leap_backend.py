import win32com.client
import json
import pandas as pd
import numpy as np
import time
"""
	THIS MODULE IS THE LEAP-BACKEND FOR FEWSIM SYSTEM
"""

input_variable = ['Activity Level','Load Shape', 'Final Energy Intensity', 'Final Energy Intensity Time Sliced',
                  'Planning Reserve Margin', 'Optimize', 'Additions to Reserves', 'Resource Imports',
                  'Resource Exports','Unmet Requirements']
def generate_LEAP_variables():
	"""
	This function extract all results values from LEAP
	:return: Structured dictionary of LEAP results value
	"""
	win32com.CoInitialize()
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	start_year = LEAP.BaseYear
	end_year = LEAP.EndYear
	# LEAP.ActiveArea = 'Internal_Linking_test'
	active_scenario = ''

	for s in LEAP.Scenarios:
		if s != 'Current Account':
			active_scenario = s
	LEAP.ActiveScenario = active_scenario
	input_variable_list = pd.read_excel('LEAP_Input_Variables.xlsx')
	input_variable_list = np.array(input_variable_list['variable_name'])
	LEAP_input = []
	LEAP_output = []
	for b in LEAP.Branches:
		print('\n')
		print(b.FullName)
		print(b.Name)
		for v in LEAP.Branch(b.FullName).Variables:
			value = []
			if v != None:
				for y in range(start_year, end_year+1):
					value.append(v.Value(y))
				path = path_parser(b.FullName)
				path.append(v.name)
				node = {
					'name':v.name,
					'fullname': b.FullName,
					'path': path,
					'parent': path[-2] if len(path)>1 else 'null',
					'value': value
				}
			if v.name in input_variable_list:
				LEAP_input = tree_insert_node(path, node, LEAP_input)
			else:
				LEAP_output = tree_insert_node(path, node, LEAP_output)
			print(node)
			# LEAP_input = tree_insert_node(path_parser(v.FullName), node, LEAP_input)
		# if v.IsResultVariable == True:
		# 	print(v.name)
		# print(type(LEAP.Branch(b.FullName)))
		# path = b.FullName + ":" +v.name
		# print(LEAP.ResultValue(path, 2002, 1, 'Linkage', 2002,12, 'Total'))
	# print(LEAP.Branch("\Demand\Water unrelated\Per capita demand").Variable('Energy Demand Final Units').Value(2002, 'MWH'))
	print(LEAP_input)
	with open('LEAP_variables.json', 'w') as outfile:
		json.dump([{'name': 'leap-input', 'children': LEAP_input},
		           {'name': 'leap-output', 'children': LEAP_output}], outfile)
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
	"""
	This module is used to query the variable tree.
	:param path_key: The path for the node to be queried.
	:param tree: The tree to be query from.
	:return: The node value.
	"""
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
	"""
	This module is used to insert a node to the tree.
	:param path_key: The path of a node to be inserted.
	:param node: The node to be inserted.
	:param tree: The tree to which the node is inserted.
	:return: The tree with new nodes inserted.
	"""
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


def get_LEAP_variables_from_file(file_path):
	"""
	This module grabs the list of LEAP variables and their paths from the stored local JSON file
	:param file_path: The path of the local file
	:return: input_list of all the LEAP inputs
	"""
	with open(file_path) as f:
		variables = json.load(f)
	input_list = []
	input_list = expand_tree(variables, input_list)
	print(len(input_list))
	return input_list

def get_LEAP_variables_tree(file_path):
	with open(file_path) as f:
		variables = json.load(f)
	return variables

# generate_LEAP_variables()
# start_time = time.time()
# generate_LEAP_variables()
# elapsed_time = time.time() - start_time
# print('Extraction of all LEAP variables takes: ',elapsed_time, ' s')
# get_LEAP_variables_from_file('LEAP_variables.json')
# path_key = ['Top Level: A', 'Level 2: A', 'Son of A']
#
# node = {"name": path_key[-1],
# 		"parent": path_key[-2],
# 		"children": [1,2,3]}



# get_LEAP_Variables()
# LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
# for v in LEAP.Branch('Key\\Population').Variables:
# 	print(v.ScaleUnit)
# LEAP.Branch('Key\\Population').Variables