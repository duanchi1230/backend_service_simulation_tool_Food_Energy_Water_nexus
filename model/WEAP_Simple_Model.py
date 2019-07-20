import win32com.client

def get_WEAP_flow_value():
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	start_year = WEAP.BaseYear
	end_year = WEAP.EndYear
	area = ['Internal_linking_test', 'WEAP_Test_Area']
	WEAP.ActiveArea = area[0]
	link = []
	path = []
	node = ''
	switch = False
	for branch in WEAP.Branches:
		name = branch.Name
		if name == 'Return Flows':
			break
		if switch == True:
			if name[0:2] == 'to':
				node = name
			if node != name:
				path.append({'demand': str(node), 'source': str(name), 'path': (str(node)+'\\'+str(name)) })
		if name == 'Transmission Links':
			switch = True

	flow = {}
	for s in WEAP.Scenarios:
		output = []
		if str(s) != 'Current Accounts':
			for p in path:
				item = {}
				value_year = []
				for year in range(start_year+1,end_year+1):
					value = WEAP.ResultValue('\Supply and Resources\Transmission Links\\' + p['path'] + ':Flow[m^3]', year, 1, str(s),
					                  year, 12, 'Total')
					value_year.append(value)

				item['name'] = p['source'][5:] + ' ' +p['demand']
				item['site'] = p['demand'][3:]
				item['source'] = p['source']
				item['value'] = value_year
				item['format'] = 'series'
				output.append(item)
			flow[str(s)] = output
	timeRange = [start_year + 1, end_year]
	win32com.CoUninitialize()
	return flow, timeRange
flow, timeRange= get_WEAP_flow_value()
print(flow)
print(flow[list(flow.keys())[0]])









class Flow:
	def __init__(self):
		self.value = {
			'\\from CAPWithdral': 0,
			'\\from GW': 0,
			'\\from GW_SRP': 0,
			'\\from SRPwithdral': 0,
			'\\from WWTP': 0
		}

# def get_WEAP_flow_value():
# 	win32com.CoInitialize()
# 	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
# 	# WEAP.AutoCalc = 'TRUE'
# 	flow = {}
# 	node = ['\\to Municipal', '\\to Agriculture', '\\to Agriculture2',
# 	        '\\to Industrial', '\\to PowerPlant', '\\to SRP_GW', '\\to Indian']
# 	link = ['\\from CAPWithdral', '\\from GW', '\\from SRPwithdral', '\\from WWTP']
# 	for j in range(3):
# 		flow_senario = {}
# 		for year in range(start_year, end_year):
# 			Municipal_flow = Flow()
# 			Agriculture_flow = Flow()
# 			Agriculture2_flow = Flow()
# 			Industrial_flow = Flow()
# 			PowerPlant_flow = Flow()
# 			Indian_flow = Flow()
# 			for i in range(len(link)):
# 				Municipal_flow.value[link[i]] = WEAP.ResultValue(
# 					'\Supply and Resources\Transmission Links' + node[0] + link[i] + ':Flow[m^3]', year, 1,
# 					scenarios[j], year, 12, 'Total')
# 				# print(link[i], 'Municipal', Municipal_flow.value[link[i]], year)
# 				Agriculture_flow.value[link[i]] = WEAP.ResultValue(
# 					'\Supply and Resources\Transmission Links' + node[1] + link[i] + ':Flow[m^3]', year, 1,
# 					scenarios[j], year, 12, 'Total')
# 				Agriculture2_flow.value[link[i]] = WEAP.ResultValue(
# 					'\Supply and Resources\Transmission Links' + node[2] + link[i] + ':Flow[m^3]', year, 1,
# 					scenarios[j], year, 12, 'Total')
# 				Industrial_flow.value[link[i]] = WEAP.ResultValue(
# 					'\Supply and Resources\Transmission Links' + node[3] + link[i] + ':Flow[m^3]', year, 1,
# 					scenarios[j], year, 12, 'Total')
#
# 			PowerPlant_flow.value['\\from WWTP'] = WEAP.ResultValue(
# 				'\Supply and Resources\Transmission Links' + node[4] + link[3] + ':Flow[m^3]', year, 1, scenarios[j],
# 				year, 12, 'Total')
# 			PowerPlant_flow.value['\\from GW'] = WEAP.ResultValue(
# 				'\Supply and Resources\Transmission Links' + node[4] + link[1] + ':Flow[m^3]', year, 1, scenarios[j],
# 				year, 12, 'Total')
# 			Indian_flow.value['\\from CAPWithdral'] = WEAP.ResultValue(
# 				'\Supply and Resources\Transmission Links' + node[6] + link[0] + ':Flow[m^3]', year, 1, scenarios[j],
# 				year, 12, 'Total')
# 			Indian_flow.value['\\from GW'] = WEAP.ResultValue(
# 				'\Supply and Resources\Transmission Links' + node[6] + link[1] + ':Flow[m^3]', year, 1, scenarios[j],
# 				year, 12, 'Total')
# 			Indian_flow.value['\\from SRPwithdral'] = WEAP.ResultValue(
# 				'\Supply and Resources\Transmission Links' + node[6] + link[2] + ':Flow[m^3]', year, 1, scenarios[j],
# 				year, 12, 'Total')
# 			# print(Municipal_flow.value)
# 			# print(Agriculture_flow.value)
# 			# print(Agriculture2_flow.value)
# 			# print(Industrial_flow.value)
# 			# print(PowerPlant_flow.value)
# 			# print(Indian_flow.value)
# 			flow_year = {}
# 			flow_year['Municipal'] = Municipal_flow.value
# 			flow_year['Agriculture'] = Agriculture_flow.value
# 			flow_year['Agriculture2'] = Agriculture2_flow.value
# 			flow_year['Industrial'] = Industrial_flow.value
# 			flow_year['PowerPlant'] = PowerPlant_flow.value
# 			flow_year['Indian'] = Indian_flow.value
# 			flow_senario[str(year)] = flow_year
# 		flow[scenarios[j]] = flow_senario
#
# 	win32com.CoUninitialize()
# 	sites = ['Municipal', 'Agriculture', 'Agriculture2', 'Industrial', 'PowerPlant', 'Indian']
# 	source = {'\\from CAPWithdral': 'CAP to', '\\from GW': 'GW to', '\\from SRPwithdral': 'SRP to',
# 	          '\\from WWTP': 'WWTP to'}
# 	#######################################################################################################
# 	'''
# 	Sort the flow value in the following data structure
# 		{
# 			'scenario':{
# 				'name': '',
# 				'value': [],
# 				'delta_to_reference': [],
# 				'site': '',
# 				'source': ''
# 			}
#
# 		}
# 	'''
#
# 	value = {}
# 	for s in scenarios:
# 		value_year = []
# 		for site in sites:
# 			for l in link:
# 				var = []
# 				for y in range(start_year, end_year):
# 					var.append(flow[s][str(y)][site][l])
# 				# print(flow[s][str(y)][site][l])
# 				# print(var)
# 				value_year.append(
# 					{'name': source[l] + ' ' + site, 'value': var, 'site': site, 'source': source[l],
# 					 'format': 'series'})
# 		value[s] = value_year
# 		value['timeRange'] = [start_year, end_year]
# 	# Add percentage change compared to 'Reference' scenario
# 	for s in scenarios:
# 		for i in range(len(value[s])):
# 			v = []
# 			for j in range(len(value[s][i]['value'])):
# 				v.append('{:.1%}'.format((value[s][i]['value'][j] - value[scenarios[0]][i]['value'][j]) / (
# 						value[scenarios[0]][i]['value'][j] + 0.1)))
# 			value[s][i]['delta_to_reference'] = v
# 	# print(value[scenarios[2]][2]['delta_to_reference'])
# 	return value
