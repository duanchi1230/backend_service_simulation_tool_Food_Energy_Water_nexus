import win32com.client
import json
import matplotlib.pyplot as plt

"""
	Module Name: WEAP and LEAP time step control
	Purpose: This module is used to test controlling WEAP (include MABIA economical model and LEAP coupled model 
			to run by monthly steps
	Status: 1. Still testing WEAP and LEAP coupled model
			2. MABIA economical model not included yet
"""
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')


def get_WEAP_flow_value():
	"""
	This module extracts the result from WEAP and
	:return: Flow(value) and timeRange
	"""
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
	print('flow_by_step', flow)
	return flow, timeRange


def get_LEAP_value():
	"""
	This part is still under development
	:return: LEAP_Results
	"""
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
	start_year = LEAP.BaseYear
	end_year = LEAP.EndYear
	LEAP_Result = []
	demand = ['Per capita demand', 'CAP pumping', 'WTP', 'WWTP']
	transformation = ['Power1', 'Power2']
	resources = ['Nuclear', 'Natural Gas', 'Electricity']
	path = {}
	for d in demand:
		if d == 'Per capita demand':
			path[d] = {'name': d, 'path': 'Demand\\Water unrelated\\' + d, 'variable': 'Energy Demand Final Units',
			           'unit': 'MWH'}
		else:
			path[d] = {'name': d, 'path': 'Demand\\Water related\\' + d, 'variable': 'Energy Demand Final Units',
			           'unit': 'MWH'}
	for t in transformation:
		path[t] = {'name': t, 'path': 'Transformation\\Electricity generation\\Processes\\' + t,
		           'variable': 'Average Power Dispatched', 'unit': 'MW'}
	for r in resources:
		if r == 'Electricity':
			path[r] = {'name': r, 'path': 'Resources\\Secondary\\' + r, 'variable': 'Primary Requirements', 'unit':'MWH'}
		else:
			path[r] = {'name': r, 'path': 'Resources\\Primary\\' + r, 'variable': 'Primary Requirements', 'unit':'MWH'}

	for v in path.values():
		print(v)
		step_result = {}
		step_result['name'] = v['name']
		step_result['variable'] = v['variable']
		step_result['value'] = []
		for y in range(start_year + 1, end_year + 1):
			step_result['value'].append(
				LEAP.Branch(v['path']).Variable(v['variable']).Value(y, v['unit']))
		step_result['unit'] = v['unit']
		LEAP_Result.append(step_result)
	print(LEAP_Result)
	timeRange = [start_year + 1, end_year]
	# LEAP.Branch('Resources\\Primary\\Nuclear').Variable('Primary Requirements').Value(2002)
	return LEAP_Result, timeRange


def iterate_by_month():
	"""
	This module runs WEAP and LEAP by steps and is still under development.
	:return: WEAP_Result, LEAP_Result
	"""
	year = [2001, 2005]
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002
	LEAP.EndYear = 2002
	LEAP_Result = {}
	LEAP_Population_Growth = 0.03
	LEAP_Population = 1000000

	WEAP.BaseYear = 2001
	WEAP.EndYear = 2002
	WEAP_Result = {}
	WEAP_Population_Growth = 0.03
	WEAP_Population = 100000
	WEAP.ActiveScenario = "Current Accounts"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 100000
	WEAP.BranchVariable('\Supply and Resources\Groundwater\Agricultural Groundwater:Initial Storage').Expression = 10000
	WEAP.ActiveScenario = "Linkage"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 'Growth(0.03)'

	for y in range(year[0], year[1]):
		WEAP.BaseYear = y
		WEAP.EndYear = y + 1
		LEAP.EndYear = y + 1
		LEAP.FirstScenarioYear = y + 1
		LEAP.BaseYear = y
		WEAP.Calculate()
		flow, WEAP_timeRange = get_WEAP_flow_value()
		LEAP.Calculate()
		value, LEAP_timeRange = get_LEAP_value()

		WEAP.ActiveScenario = "Current Accounts"
		WEAP.BranchVariable('\Key Assumptions\Population').Expression = WEAP_Population * (
				1 + WEAP_Population_Growth) ** (y - year[0] + 1)
		print('Population', WEAP.BranchVariable('\Key Assumptions\Population').Expression)
		WEAP.ActiveScenario = "Current Accounts"
		WEAP.BranchVariable(
			'\Supply and Resources\Groundwater\Agricultural Groundwater:Initial Storage').Expression = WEAP.ResultValue(
			'\Supply and Resources\Groundwater\Agricultural Groundwater: Storage',
			y + 1, 12, 'Linkage', y + 1, 12, 'Total') / 1000000
		LEAP.Branch('\Key Assumptions\Population').Variable().Expression = LEAP_Population * (
				1 + LEAP_Population_Growth) ** (y - year[0] + 1)
		print('LEAP population', LEAP.Branch('\Key Assumptions\Population').Variable())
		# print('Groundwater Storage', WEAP.ResultValue(
		# 	'\Supply and Resources\Groundwater\Agricultural Groundwater: Storage',
		# 	y+1, 12, 'Linkage', y+1, 12, 'Total')/1000000)

		WEAP_Result[str(y + 1)] = flow
		LEAP_Result[str(y + 1)] = value
		print(flow)
		print(y + 1)

	return WEAP_Result, LEAP_Result


def run_fulltime_WEAP():
	"""
	This module runs the WEAP on the full time span (bulk_run)
	:return:
	"""
	WEAP.ActiveScenario = "Current Accounts"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 100000
	WEAP.BranchVariable(
		'\Supply and Resources\Groundwater\Agricultural Groundwater:Initial Storage').Expression = 100000
	WEAP.ActiveScenario = "Linkage"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 'Growth(0.03)'
	WEAP.BaseYear = 2001
	WEAP.EndYear = 2005

	LEAP.ActiveScenario = 'Current Accounts'
	LEAP.Branch('\Key Assumptions\Population').Variable().Expression = 1000000
	LEAP.ActiveScenario = 'Linkage'
	LEAP.Branch('\Key Assumptions\Population').Variable().Expression = 'Growth(0.03)'
	LEAP.EndYear = 2005
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002

	WEAP.Calculate()
	flow, timeRange = get_WEAP_flow_value()
	return flow, timeRange


def run_fulltime_LEAP():
	"""
	This module runs the WEAP on the full time span (bulk_run)
	:return:
	"""
	WEAP.ActiveScenario = "Current Accounts"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 100000
	WEAP.BranchVariable(
		'\Supply and Resources\Groundwater\Agricultural Groundwater:Initial Storage').Expression = 100000
	WEAP.ActiveScenario = "Linkage"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 'Growth(0.03)'
	WEAP.BaseYear = 2001
	WEAP.EndYear = 2005

	LEAP.ActiveScenario = 'Current Accounts'
	LEAP.Branch('\Key Assumptions\Population').Variable().Expression = 1000000
	LEAP.ActiveScenario = 'Linkage'
	LEAP.Branch('\Key Assumptions\Population').Variable().Expression = 'Growth(0.03)'
	LEAP.EndYear = 2005
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002

	LEAP.Calculate()
	value, timeRange = get_LEAP_value()
	return value, timeRange


def reformat_WEAP_Result(WEAP_Result):
	"""
	This module reforms the step run result into the same format as bulk run to facilitate the compare module.
	:param WEAP_Result: Step run result
	:return: Reformated step run result same as bulk run
	"""
	result = WEAP_Result[list(WEAP_Result.keys())[0]]
	for k in WEAP_Result:
		if k != list(WEAP_Result.keys())[0]:
			for s in WEAP_Result[k]:
				for c in WEAP_Result[k][s]:
					for var in result[s]:
						if var['name'] == c['name']:
							var['value'].append(c['value'][0])
	print('2', result['Linkage'])
	return result['Linkage']


def reformat_LEAP_Result(LEAP_Result):
	"""
	This module reforms the step run result into the same format as bulk run to facilitate the compare module.
	:param LEAP_Result: Step run result
	:return: Reformated step run result same as bulk run
	"""
	result = LEAP_Result['2002']
	for k in LEAP_Result:
		if k != list(LEAP_Result.keys())[0]:
			for s in LEAP_Result[k]:
				for var in result:
					if var['name'] == s['name']:
						var['value'].append(s['value'][0])
	print(result)
	return result


def compare_result_WEAP(flow, WEAP_Result):
	"""
	This module compare the bulk run and step run result from WEAP.
	:param flow: The bulk run result from WEAP.
	:param WEAP_Result: The step run result from WEAP.
	:return: The plots of comparing the bulk run and step run results from WEAP.
	"""
	# print(flow[list(flow.keys())[0]])
	year = ['2002', '2003', '2004', '2005']

	for i in range(len(flow[list(flow.keys())[0]])):
		fig = plt.figure(0)
		plt.plot(year, flow[list(flow.keys())[0]][i]['value'], color='blue',
		         label='Bulk_Run ' + flow[list(flow.keys())[0]][i]['name'])
		plt.plot(year, WEAP_Result[i]['value'], color='orange', label='Step_Run ' + WEAP_Result[i]['name'], alpha=0.7,
		         linewidth=5)
		plt.xlabel('Scenario Year')
		plt.ylabel('Flow Value: M^3')
		plt.title('WEAP Result Value: ' + WEAP_Result[i]['name'])
		plt.legend(loc=1)
		bottom, top = plt.ylim()
		plt.ylim(bottom / 3, top * 1.2)
		plt.gcf().autofmt_xdate()
		# print(flow[list(flow.keys())[0]])
		fig.show()


def compare_result_LEAP(value, LEAP_Result):
	"""
	This module compares the bulk run and step run results for LEAP.
	:param value: The bulk run result from LEAP
	:param LEAP_Result: The step run result from LEAP
	:return: The plots of comparing the bulk run and step run results from LEAP
	"""
	year = ['2002', '2003', '2004', '2005']

	for i in range(len(value)):
		fig = plt.figure(0)
		plt.plot(year, value[i]['value'], color='blue', label='Bulk_Run ' + value[i]['name'])
		plt.plot(year, LEAP_Result[i]['value'], color='orange', label='Step_Run ' + LEAP_Result[i]['name'], alpha=0.7,
		         linewidth=5)
		plt.xlabel('Scenario Year (Unconstrained Source and Unconstrained Link)')
		plt.ylabel('Unit: ' + value[i]['unit'])
		plt.title('LEAP Result Value: ' + LEAP_Result[i]['name'] + ' ' +LEAP_Result[i]['variable'])
		plt.legend(loc=1)
		bottom, top = plt.ylim()
		plt.ylim(bottom / 3, top * 1.2)
		plt.gcf().autofmt_xdate()
		# print(flow[list(flow.keys())[0]])
		fig.show()
	pass


# get_LEAP_value()
WEAP_Result, LEAP_Result = iterate_by_month()
print(LEAP_Result)
with open('WEAP_Result.json', 'w') as f:
	json.dump(WEAP_Result, f)
# with open('LEAP_Result.json', 'w') as f:
# 	json.dump(LEAP_Result, f)
#
# with open('LEAP_Result.json') as wp:
# 	LEAP_Result = json.load(wp)
# LEAP_Result = reformat_LEAP_Result(LEAP_Result)
# value, timeRange = run_fulltime_LEAP()
# compare_result_LEAP(value, LEAP_Result)
# print(timeRange, value)


with open('WEAP_Result.json') as wp:
	WEAP_Result = json.load(wp)
WEAP_Result = reformat_WEAP_Result(WEAP_Result)
print(WEAP_Result)
flow, timeRange = run_fulltime_WEAP()
print(flow)
compare_result_WEAP(flow, WEAP_Result)