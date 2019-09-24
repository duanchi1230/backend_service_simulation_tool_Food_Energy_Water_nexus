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
	print(LEAP.Branch('Transformation\Electricity generation\Processes\Power2').Variable(
		'Average Power Dispatched').Value(2002))
	print(LEAP.Branch('Resources\Primary\\Natural Gas').Variable(
		'Imports').Value(2002))
	print(LEAP.Branch('Demand\\Water unrelated\\Per capita demand').Variable(
		'Energy Demand Final Units').Value(2002))
	demand = ['Per capita demand', 'CAP pumping', 'WTP', 'WWTP']
	transformation = ['Power1', 'Power2']
	resources = ['Nuclear', 'Natural Gas', 'Electricity']
	for b in LEAP.Branches:
		print(b.name)
	path = {}
	for d in demand:
		if d=='Per capita demand':
			path[d] = {'name': d, 'path': 'Demand\\Water unrelated\\' + d, 'variable': 'Energy Demand Final Units'}
		else: path[d] = {'name': d, 'path': 'Demand\\Water related\\' + d, 'variable': 'Energy Demand Final Units'}
	for t in transformation:
		path[t] = {'name': t, 'path': 'Transformation\\Electricity generation\\Processes\\' + t, 'variable': 'Average Power Dispatched'}
	for r in resources:
		if r =='Electricity':
			path[r] = {'name': r, 'path': 'Resources\\Secondary\\' + r, 'variable': 'Primary Supply'}
		else: path[r] = {'name': r, 'path': 'Resources\\Primary\\' + r, 'variable': 'Primary Supply'}

	for v in path.values():
		print(v)
	# for y in range(start_year + 1, end_year + 1):
	# 	print(y)
	# 	step_result = {}
	# 	step_result['Energy Demand Final Units'] = [
	# 	LEAP.Branch('Demand\\Water unrelated\\' + d).Variable(
	# 		'Energy Demand Final Units').Value(2002)]
	# print(LEAP_Result)


def iterate_by_month():
	"""
	This module runs WEAP and LEAP by steps and is still under development.
	:return: WEAP_Result, LEAP_Result
	"""
	year = [2001, 2005]
	LEAP.BaseYear = 2001
	LEAP.FirstScenarioYear = 2002
	LEAP.EndYear = 2002
	WEAP.BaseYear = 2001
	WEAP.EndYear = 2002
	WEAP_Result = {}
	WEAP_Population_Growth = 0.03
	WEAP_Population = 100000
	LEAP_Population_Growth = 0.03
	LEAP_Population = 1000000

	WEAP.ActiveScenario = "Current Accounts"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 100000
	WEAP.BranchVariable('\Supply and Resources\Groundwater\Agricultural Groundwater:Initial Storage').Expression = 10000
	WEAP.ActiveScenario = "Linkage"
	WEAP.BranchVariable('\Key Assumptions\Population').Expression = 'Growth(0.00)'

	for y in range(year[0], year[1]):
		WEAP.BaseYear = y
		WEAP.EndYear = y + 1
		LEAP.EndYear = y + 1
		LEAP.FirstScenarioYear = y + 1
		LEAP.BaseYear = y
		WEAP.ActiveScenario = "Current Accounts"
		WEAP.BranchVariable('\Key Assumptions\Population').Expression = WEAP_Population * (
				1 + WEAP_Population_Growth) ** (y - year[0] + 1)
		print('Population', WEAP.BranchVariable('\Key Assumptions\Population').Expression)
		WEAP.Calculate()
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
		flow, timeRange = get_WEAP_flow_value()
		WEAP_Result[str(y + 1)] = flow
		print(flow)
		print(y + 1)

	return WEAP_Result


def run_full_time():
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


def compare_result_WEAP(flow, WEAP_Result):
	"""
	This module compare the bulk run and step run result.
	:param flow: The bulk run result
	:param WEAP_Result: The step run result
	:return: The plots of comparing the bulk run and step run results
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


def compare_result_LEAP():
	pass


# WEAP_Result = iterate_by_month()
# with open('WEAP_Result.json', 'w') as f:
# 	json.dump(WEAP_Result, f)

# with open('WEAP_Result.json') as wp:
# 	WEAP_Result = json.load(wp)
# WEAP_Result = reformat_WEAP_Result(WEAP_Result)
# print(WEAP_Result)
# flow, timeRange = run_full_time()
# print(flow)
# compare_result_WEAP(flow, WEAP_Result)
#
get_LEAP_value()
