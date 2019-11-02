import win32com.client
import json

"""
	THIS MODULE IS THE LEAP-BACKEND FOR FEWSIM SYSTEM
"""


def get_LEAP_Outputs():
	"""
	This function extract all results values from LEAP
	:return: Structured dictionary of LEAP results value
	"""
	win32com.CoInitialize()
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	start_year = LEAP.BaseYear
	end_year = LEAP.EndYear
	LEAP.ActiveArea = 'Internal_Linking_test'
	active_scenario = ''

	for s in LEAP.Scenarios:
		if s != 'Current Account':
			active_scenario = s
	LEAP.ActiveScenario = active_scenario

	LEAP_input = []
	LEAP_output = []
	for b in LEAP.Branches:
		LEAP.Branch(b.FullName)
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
			LEAP_input = tree_insert_node(path, node, LEAP_input)
			print(node)
			# LEAP_input = tree_insert_node(path_parser(v.FullName), node, LEAP_input)
		# if v.IsResultVariable == True:
		# 	print(v.name)
		# print(type(LEAP.Branch(b.FullName)))
		# path = b.FullName + ":" +v.name
		# print(LEAP.ResultValue(path, 2002, 1, 'Linkage', 2002,12, 'Total'))
	# print(LEAP.Branch("\Demand\Water unrelated\Per capita demand").Variable('Energy Demand Final Units').Value(2002, 'MWH'))
	print(LEAP_input)
	with open('LEAP_input.txt', 'w') as outfile:
		json.dump(LEAP_input, outfile)
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

path_key = ['Top Level: A', 'Level 2: A', 'Son of A']

node = {"name": path_key[-1],
		"parent": path_key[-2],
		"children": [1,2,3]}

# WEAP_tree = []
# WEAP_tree = tree_insert_node(path_key, node, WEAP_tree)
# print(WEAP_tree)
# print(tree_find_key(path_key, WEAP_tree))

get_LEAP_Outputs()

