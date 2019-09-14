import win32com.client

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
	win32com.CoInitialize()
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
	win32com.CoUninitialize()
	return flow, timeRange


def iterate_by_month():
	year = [2001, 2005]
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002
	LEAP.EndYear = 2002
	WEAP.BaseYear = 2001
	WEAP.EndYear = 2002
	for y in range(year[0], year[1]):
		# WEAP.BaseYear = y
		# WEAP.EndYear = y+1
		LEAP.EndYear = y + 1
		LEAP.FirstScenarioYear = y+1
		LEAP.BaseYear = y
		print('1',LEAP.BaseYear)
		print('2',LEAP.FirstScenarioYear)
		print('3',LEAP.EndYear)
		LEAP.Calculate()
		# print(WEAP.BaseYear)
		# print(WEAP.EndYear)
		# WEAP.Calculate()
		# print(get_WEAP_flow_value())
		v2 = WEAP.ResultValue(
			'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 1:Total Node Outflow[m^3]', y,
			1,
			'Linkage', y, 12, 'Total')
		print(v2)


iterate_by_month()
