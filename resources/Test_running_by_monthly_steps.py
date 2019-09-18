import win32com.client
import json
import matplotlib.pyplot as plt
'''
	Module Name: WEAP and LEAP time step control
	Purpose: This module is used to test controlling WEAP (include MABIA economical model and LEAP coupled model 
			to run by monthly steps
	Status: 1. Still testing WEAP and LEAP coupled model
			2. MABIA economical model not included yet
'''
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')


def get_WEAP_flow_value():
	### Initialize win32com object###
	# win32com.CoInitialize()
	### Initialize WEAP application communication port###
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear

	area = ['Internal_linking_test', 'WEAP_Test_Area', 'Internal_Linking_test_das']
	WEAP.ActiveArea = area[2]
	link = []
	path = []
	node = ''
	switch = False
	### Extract the pathes for Transmission Links ###
	for branch in WEAP.Branches:
		name = branch.Name
		if name == 'Runoff and Infiltration':
			break
		if switch == True:
			if name[0:2] == 'to':
				node = name
			if node != name:
				path.append({'demand': str(node), 'source': str(name), 'path': (str(node) + '\\' + str(name))})
		if name == 'Transmission Links':
			switch = True
	### Extract results for each Transmission Link ###
	flow = {}
	for s in WEAP.Scenarios:
		output = []
		if str(s) != 'Current Accounts':
			for p in path:
				item = {}
				value_year = []
				for year in range(start_year + 1, end_year + 1):
					value = WEAP.ResultValue('\Supply and Resources\Transmission Links\\' + p['path'] + ':Flow[m^3]',
					                         year, 1, str(s),
					                         year, 12, 'Total')
					value_year.append(value)

				item['name'] = p['source'][5:] + ' ' + p['demand']
				item['site'] = p['demand'][3:]
				item['source'] = p['source']
				item['value'] = value_year
				item['format'] = 'series'
				output.append(item)
			flow[str(s)] = output
	timeRange = [start_year + 1, end_year]
	### Uninitialize the win32com object ###
	# win32com.CoUninitialize()

	return flow, timeRange


def iterate_by_month():
	year = [2001, 2005]
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002
	LEAP.EndYear = 2002
	WEAP.BaseYear = 2001
	WEAP.EndYear = 2002
	WEAP_Result ={}
	WEAP_Population_Growth = 0.03
	for y in range(year[0], year[1]):
		WEAP.BaseYear = y
		WEAP.EndYear = y+1
		LEAP.EndYear = y + 1
		LEAP.FirstScenarioYear = y+1
		LEAP.BaseYear = y
		print('1',LEAP.BaseYear)
		print('2',LEAP.FirstScenarioYear)
		print('3',LEAP.EndYear)
		LEAP.Calculate()
		# WEAP.Calculate()
		flow, timeRange = get_WEAP_flow_value()
		WEAP_Result[str(y+1)] = flow
		print(flow)
		print(y+1)
		# print(WEAP_Result)
		# v2 = WEAP.ResultValue(
		# 	'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 1:Total Node Outflow[m^3]', y,
		# 	1,
		# 	'Linkage', y, 12, 'Total')
		# print(v2)
	return WEAP_Result

def reformat_WEAP_Result(WEAP_Result):
	result = WEAP_Result[list(WEAP_Result.keys())[0]]
	for k in WEAP_Result:
		print(k)
		if k !=list(WEAP_Result.keys())[0]:
			for s in WEAP_Result[k]:
				for c in WEAP_Result[k][s]:
					for var in result[s]:
						if var['name']==c['name']:
							var['value'].append(c['value'][0])
	print('2', result[s])
	return result[s]

def compare_result(flow, WEAP_Result):
	print(flow[list(flow.keys())[0]])

	for i in range(len(flow[list(flow.keys())[0]])):
		fig = plt.figure(0)
		plt.plot(range(2002,2006), flow[list(flow.keys())[0]][i]['value'])
		plt.plot(range(2002,2006), WEAP_Result[i]['value'])
		fig.show()



# WEAP_Result = iterate_by_month()
# with open('WEAP_Result.json', 'w') as f:
# 	json.dump(WEAP_Result, f)
with open('WEAP_Result.json') as wp:
	WEAP_Result = json.load(wp)

WEAP_Result = reformat_WEAP_Result(WEAP_Result)
flow, timeRange = get_WEAP_flow_value()

compare_result(flow, WEAP_Result)

