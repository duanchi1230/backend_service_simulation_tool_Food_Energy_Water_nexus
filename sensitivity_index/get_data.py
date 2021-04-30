import numpy as np
import pandas as pd
import jenkspy

def get_data():
	LEAP_SI = pd.read_excel("./sensitivity_index/data/Results_SI_LEAP_138.xlsx", skiprows=[0], header=None)
	WEAP_SI = pd.read_excel("./sensitivity_index/data/Results_SI_WEAP_standard23.xlsx", skiprows=[0], header=None)
	leap_input = pd.read_csv("./sensitivity_index/data/SA_inputs_LEAP_Mar12_2020.csv")
	leap_output = pd.read_csv("./sensitivity_index/data/SA_outputs_LEAP_Mar12_2020.csv")
	
	weap_input = pd.read_csv("./sensitivity_index/data/SA_inputs_WEAP_6May2020.csv")
	weap_output = pd.read_csv("./sensitivity_index/data/SA_outputs_WEAP_6May2020.csv")
	print(LEAP_SI)
	# print(WEAP_SI)

	# print(leap_input)
	# print(leap_output)
	#
	# print(weap_input)
	# print(weap_output)
	
	LEAP_SI_matrix = []
	WEAP_SI_matrix = []
	i = 0
	while i < len(LEAP_SI.columns):
		LEAP_SI_matrix.append(LEAP_SI[i + 2].tolist())
		i += 3
	LEAP_SI_matrix = pd.DataFrame(LEAP_SI_matrix, columns=LEAP_SI[1])
	print(LEAP_SI_matrix)
	i = 0
	while i < len(WEAP_SI.columns):
		WEAP_SI_matrix.append(WEAP_SI[i + 2].tolist())
		i += 3
	WEAP_SI_matrix = pd.DataFrame(WEAP_SI_matrix, columns=WEAP_SI[1])
	node = []
	for i in range(len(leap_input)):
		node.append({"id":leap_input.iloc[i]["branchToChange"] +":"+ leap_input.iloc[i]["variableToChange"], "group": "leap-input"})
	for i in range(len(leap_output)):
		node.append({"id":leap_output.iloc[i]["branchToSave"] +":"+  leap_output.iloc[i]["variable"], "group": "leap-output"})
	for i in range(len(weap_input)):
		node.append({"id":weap_input.iloc[i]["branchToChange"] +":"+  weap_input.iloc[i]["variableToChange"], "group": "weap-input"})
	for i in range(len(weap_output)):
		node.append({"id":weap_output.iloc[i]["branchesToSave"] +":"+  weap_output.iloc[i]["variablesToSave"], "group": "weap-output"})
	
	link = []
	index_value = []
	for i in range(len(LEAP_SI_matrix)):
		for j in LEAP_SI_matrix.columns:
			# print(i,j)
			link.append({"source": leap_input.iloc[i]["branchToChange"] +":"+  leap_input.iloc[i]["variableToChange"],
			             "target": leap_output.iloc[j]["branchToSave"] +":"+  leap_output.iloc[j]["variable"], "value": LEAP_SI_matrix.iloc[i][j], "type": "leap"})
			index_value.append(abs(LEAP_SI_matrix.iloc[i][j]))
	for i in range(len(WEAP_SI_matrix)):
		for j in WEAP_SI_matrix.columns:
			link.append({"source": weap_input.iloc[i]["branchToChange"] +":"+  weap_input.iloc[i]["variableToChange"],
			             "target": weap_output.iloc[j]["branchesToSave"] +":"+  weap_output.iloc[j]["variablesToSave"], "value": WEAP_SI_matrix.iloc[i][j], "type": "weap"})
			index_value.append(abs(WEAP_SI_matrix.iloc[i][j]))
	# for l in link:
	# 	print(l)
	# for n in node:
	# 	print(n)
	bounds = jenkspy.jenks_breaks(index_value, nb_class=9)
	print(bounds)
	bounds.append(0.01)
	bounds.sort()
	bounds[-1] = 27.75
	bin_data = np.histogram(index_value, bins=np.around(bounds, decimals=2))
	print(bin_data[0], bin_data[1])
	return node, link, {"quantity": np.log(bin_data[0]).tolist(), "bin": bin_data[1].tolist()}

def get_data_coupled():
	LEAP_SI = pd.read_excel("./sensitivity_index/data/Results_SI_LEAP_138.xlsx", skiprows=[0], header=None)
	WEAP_SI = pd.read_excel("./sensitivity_index/data/Results_SI_WEAP_standard23.xlsx", skiprows=[0], header=None)
	leap_input = pd.read_csv("./sensitivity_index/data/SA_inputs_LEAP_Mar12_2020.csv")
	leap_output = pd.read_csv("./sensitivity_index/data/SA_outputs_LEAP_Mar12_2020.csv")

	weap_input = pd.read_csv("./sensitivity_index/data/SA_inputs_WEAP_6May2020.csv")
	weap_output = pd.read_csv("./sensitivity_index/data/SA_outputs_WEAP_6May2020.csv")

	coupled_SI = pd.read_excel("./sensitivity_index/data/coupled/mod1Dec.xlsx", skiprows=[0], header=None)
	coupled_input = pd.read_csv("./sensitivity_index/data/coupled/SA_parameters_Combined_19Aug2020.csv")
	coupled_output = pd.read_csv("./sensitivity_index/data/coupled/SA_outputs_Combined_15Sep2020.csv")

	print(LEAP_SI)
	# print(leap_input)
	# print(leap_output)
	#
	# print(weap_input)
	# print(weap_output)

	LEAP_SI_matrix = []
	WEAP_SI_matrix = []
	coupled_matrix = []
	i = 0
	while i < len(LEAP_SI.columns):
		LEAP_SI_matrix.append(LEAP_SI[i + 2].tolist())
		i += 3
	LEAP_SI_matrix = pd.DataFrame(LEAP_SI_matrix, columns=LEAP_SI[1])
	# print(LEAP_SI_matrix)
	i = 0
	while i < len(WEAP_SI.columns):
		WEAP_SI_matrix.append(WEAP_SI[i + 2].tolist())
		i += 3
	WEAP_SI_matrix = pd.DataFrame(WEAP_SI_matrix, columns=WEAP_SI[1])

	i = 1
	while i < len(coupled_SI.columns):
		coupled_matrix.append(coupled_SI[i + 2].tolist())
		i += 4
	coupled_matrix = pd.DataFrame(coupled_matrix, columns=coupled_SI[2])
	print(coupled_matrix)
	node_type = {"WEAP":"weap", "LEAP":"leap", "MPM":"mpm"}
	node = []
	# for i in range(len(leap_input)):
	# 	node.append({"id": leap_input.iloc[i]["branchToChange"] + ":" + leap_input.iloc[i]["variableToChange"],
	# 				 "group": "leap-input"})
	# for i in range(len(leap_output)):
	# 	node.append(
	# 		{"id": leap_output.iloc[i]["branchToSave"] + ":" + leap_output.iloc[i]["variable"], "group": "leap-output"})
	# for i in range(len(weap_input)):
	# 	node.append({"id": weap_input.iloc[i]["branchToChange"] + ":" + weap_input.iloc[i]["variableToChange"],
	# 				 "group": "weap-input"})
	# for i in range(len(weap_output)):
	# 	node.append({"id": weap_output.iloc[i]["branchesToSave"] + ":" + weap_output.iloc[i]["variablesToSave"],
	# 				 "group": "weap-output"})
	for i in range(len(coupled_input)):
		node.append({"id": coupled_input.iloc[i]["branchToChange"] + ":" + coupled_input.iloc[i]["variableToChange"],
					 "group": node_type[coupled_input.iloc[i]["type"]]+"-input"})
	for i in range(len(coupled_output)):
		node.append({"id": coupled_output.iloc[i]["branchesToSave"] + ":" + coupled_output.iloc[i]["variablesToSave"],
					 "group": node_type[coupled_output.iloc[i]["type"]]+"-output"})


	link = []
	index_value = []
	# for i in range(len(LEAP_SI_matrix)):
	# 	for j in LEAP_SI_matrix.columns:
	# 		# print(i,j)
	# 		link.append({"source": leap_input.iloc[i]["branchToChange"] + ":" + leap_input.iloc[i]["variableToChange"],
	# 					 "target": leap_output.iloc[j]["branchToSave"] + ":" + leap_output.iloc[j]["variable"],
	# 					 "value": LEAP_SI_matrix.iloc[i][j], "type": "leap"})
	# 		index_value.append(abs(LEAP_SI_matrix.iloc[i][j]))
	# for i in range(len(WEAP_SI_matrix)):
	# 	for j in WEAP_SI_matrix.columns:
	# 		link.append({"source": weap_input.iloc[i]["branchToChange"] + ":" + weap_input.iloc[i]["variableToChange"],
	# 					 "target": weap_output.iloc[j]["branchesToSave"] + ":" + weap_output.iloc[j]["variablesToSave"],
	# 					 "value": WEAP_SI_matrix.iloc[i][j], "type": "weap"})
	# 		index_value.append(abs(WEAP_SI_matrix.iloc[i][j]))
	for i in range(len(coupled_matrix)):
		for j in coupled_matrix.columns:
			link.append({"source": coupled_input.iloc[i]["branchToChange"] + ":" + coupled_input.iloc[i]["variableToChange"],
						 "target": coupled_output.iloc[j]["branchesToSave"] + ":" + coupled_output.iloc[j]["variablesToSave"],
						 "value": coupled_matrix.iloc[i][j], "type": node_type[coupled_input.iloc[i]["type"]] + "-" + node_type[coupled_output.iloc[i]["type"]]})
			index_value.append(abs(coupled_matrix.iloc[i][j]))
	# for l in link:
	# 	print(l)
	# for n in node:
	# 	print(n)
	bounds = jenkspy.jenks_breaks(index_value, nb_class=9)
	print(bounds)
	bounds.append(0.01)
	bounds.sort()
	bounds[-1] = np.ceil(bounds[-1])
	bin_data = np.histogram(index_value, bins=np.around(bounds, decimals=2))
	print(bin_data[0], bin_data[1])
	print(node,link)
	return node, link, {"quantity": np.log(bin_data[0]).tolist(), "bin": bin_data[1].tolist()}
# get_data()
