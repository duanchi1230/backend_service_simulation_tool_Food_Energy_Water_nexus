import win32com.client
import json
import numpy as np
from operator import sub

WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
WEAP.ActiveArea = 'WEAP_Test_Area'
WEAP.BaseYear = 1985
WEAP.EndYear = 2009
scenarios = ['Reference', '5% Population Growth', '10% Population Growth']
start_year = 1986
end_year = 2009


# WEAP.Calculate()

# for i in range(1, WEAP.Branch('\Demand Sites').Children.Count+1):
#     print(WEAP.Branch('\Demand Sites').Children.Item(i).Children.Item(0))

def get_WEAP_para_value(path):
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	value = {}
	for s in scenarios:
		WEAP.ActiveScenario = s
		value[s] = (WEAP.Branch(path['branch']).Variables(path['variable']).Expression)
	# WEAP.Branch(path['branch']).Variables(path['variable']).Expression = 'Growth(5%)'
	win32com.CoUninitialize()
	return value


def set_WEAP_ParaValue():
	pass


class Demand:
	def __init__(self, site):
		self.site = site

		self.state = {
			'Annual Activity Level': 0,
			'Annual Water Use Rate': 0,
			'Monthly Variation': 0,
			'Consumption': 0,
			'DSM Savings': 0,
			'DSM Cost': 0,
			'Demand Priority': 0,
		}

	def setDefault(self):
		WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
		WEAP.ActiveArea = 'WEAP_Test_Area'
		for var in self.state:
			WEAP.Branch('\Demand Sites' + self.site).Variables(var).Expression = self.state[var]
			print(self.state[var])
		return 'This has been set default!'


class River:
	def __init__(self, site):
		self.site = site

		self.state = {
			'Headflow': 0,
		}

	def setDefault(self):
		WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
		WEAP.ActiveArea = 'WEAP_Test_Area'
		for var in self.state:
			WEAP.Branch('\Supply and Resources\River' + self.site).Variables(var).Expression = self.state[var]
			print(self.state[var])


class Groundwater:

	def __init__(self, site):
		self.site = site
		self.state = {
			'Storage Capacity': 0,
			'Initial Storage': 0,
			'Maximum Withdrawal': 0,
			'Natural Recharge': 101,
		}

	def setDefault(self):
		WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
		WEAP.ActiveArea = 'WEAP_Test_Area'
		for var in self.state:
			WEAP.Branch('\Supply and Resources\Groundwater' + self.site).Variables(var).Expression = self.state[var]
			print(self.state[var])


class Flow:
	def __init__(self):
		self.value = {
			'\\from CAPWithdral': 0,
			'\\from GW': 0,
			'\\from GW_SRP': 0,
			'\\from SRPwithdral': 0,
			'\\from WWTP': 0
		}


# Define default Municipal input
Municipal = Demand('\Municipal')
Municipal.state['Annual Activity Level'] = 1855960
Municipal.state['Annual Water Use Rate'] = 0.3415
Municipal.state['Consumption'] = 87
Municipal.state['Demand Priority'] = 1

# Define default Agriculture input
Agriculture = Demand('\Agriculture')
Agriculture.state['Annual Activity Level'] = 353850.48
Agriculture.state['Annual Water Use Rate'] = 3.5767
Agriculture.state['Consumption'] = 90
Agriculture.state['Demand Priority'] = 1

# Define default Agriculture2 input
Agriculture2 = Demand('\Agriculture2')
Agriculture2.state['Annual Activity Level'] = 353850.48
Agriculture2.state['Annual Water Use Rate'] = 3.5767
Agriculture2.state['Consumption'] = 90
Agriculture2.state['Demand Priority'] = 1

# Define default Industrial input
Industrial = Demand('\Industrial')
Industrial.state['Annual Activity Level'] = 1
Industrial.state['Annual Water Use Rate'] = 73099.68
Industrial.state['Consumption'] = 50
Industrial.state['Demand Priority'] = 1

# Define default PowerPlat input
PowerPlant = Demand('\PowerPlant')
PowerPlant.state['Annual Activity Level'] = 1
PowerPlant.state['Annual Water Use Rate'] = 15567.56
PowerPlant.state['Consumption'] = 10
PowerPlant.state['Demand Priority'] = 1

# Define default Indian input
Indian = Demand('\Indian')
Indian.state['Annual Activity Level'] = 1
Indian.state['Annual Water Use Rate'] = 15567.56
Indian.state['Consumption'] = 10
Indian.state['Demand Priority'] = 1

# Define default supply River
CAP = River('CAP')
CAP.state['Headflow'] = 0
SRP = River('SRP')
SRP.state['Headflow'] = 39.11

# Define default Groundwater
GW = Groundwater('\GW')
GW.state['Storage Capacity'] = 164913
GW.state['Initial Storage'] = 85251.7
GW.state['Maximum Withdrawal'] = 575.54
GW.state['Natural Recharge'] = 101
GW_SRP = Groundwater('\GW_SRP')
GW_SRP.state['Storage Capacity'] = 164913
GW_SRP.state['Initial Storage'] = 85251.7
GW_SRP.state['Maximum Withdrawal'] = 0
GW_SRP.state['Natural Recharge'] = 101


# print(WEAP.Branch('\Demand Sites\Municipal').Variables('Annual Activity Level').Expression)
# print(WEAP.ResultValue('\Supply and Resources\Groundwater\GW:Storage', 1986, 1))
# print(WEAP.ResultValue('\Demand Sites\Municipal: Supply Delivered', 1986, 1, 'Reference'))
# print(WEAP.ResultValue('\Supply and Resources\Transmission Links\\to Municipal\\from GW:Flow[m^3]', 1986, 1, '10% Population Growth'))
# for sce in WEAP.Scenarios:
# 	print(sce)
# for brc in WEAP.Branches:
# 	print(brc.FullName)
# Extract the result for the supply links
def get_WEAP_flow_value():
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	# WEAP.AutoCalc = "TRUE"
	flow = {}
	node = ['\\to Municipal', '\\to Agriculture', '\\to Agriculture2',
	        '\\to Industrial', '\\to PowerPlant', '\\to SRP_GW', '\\to Indian']
	link = ['\\from CAPWithdral', '\\from GW', '\\from SRPwithdral', '\\from WWTP']
	for j in range(3):
		flow_senario = {}
		for year in range(start_year, end_year):
			Municipal_flow = Flow()
			Agriculture_flow = Flow()
			Agriculture2_flow = Flow()
			Industrial_flow = Flow()
			PowerPlant_flow = Flow()
			Indian_flow = Flow()
			for i in range(len(link)):
				Municipal_flow.value[link[i]] = WEAP.ResultValue(
					'\Supply and Resources\Transmission Links' + node[0] + link[i] + ':Flow[m^3]', year, 1,
					scenarios[j], year, 12, 'Total')
				# print(link[i], "Municipal", Municipal_flow.value[link[i]], year)
				Agriculture_flow.value[link[i]] = WEAP.ResultValue(
					'\Supply and Resources\Transmission Links' + node[1] + link[i] + ':Flow[m^3]', year, 1,
					scenarios[j], year, 12, 'Total')
				Agriculture2_flow.value[link[i]] = WEAP.ResultValue(
					'\Supply and Resources\Transmission Links' + node[2] + link[i] + ':Flow[m^3]', year, 1,
					scenarios[j], year, 12, 'Total')
				Industrial_flow.value[link[i]] = WEAP.ResultValue(
					'\Supply and Resources\Transmission Links' + node[3] + link[i] + ':Flow[m^3]', year, 1,
					scenarios[j], year, 12, 'Total')

			PowerPlant_flow.value['\\from WWTP'] = WEAP.ResultValue(
				'\Supply and Resources\Transmission Links' + node[4] + link[3] + ':Flow[m^3]', year, 1, scenarios[j],
				year, 12, 'Total')
			PowerPlant_flow.value['\\from GW'] = WEAP.ResultValue(
				'\Supply and Resources\Transmission Links' + node[4] + link[1] + ':Flow[m^3]', year, 1, scenarios[j],
				year, 12, 'Total')
			Indian_flow.value['\\from CAPWithdral'] = WEAP.ResultValue(
				'\Supply and Resources\Transmission Links' + node[6] + link[0] + ':Flow[m^3]', year, 1, scenarios[j],
				year, 12, 'Total')
			Indian_flow.value['\\from GW'] = WEAP.ResultValue(
				'\Supply and Resources\Transmission Links' + node[6] + link[1] + ':Flow[m^3]', year, 1, scenarios[j],
				year, 12, 'Total')
			Indian_flow.value['\\from SRPwithdral'] = WEAP.ResultValue(
				'\Supply and Resources\Transmission Links' + node[6] + link[2] + ':Flow[m^3]', year, 1, scenarios[j],
				year, 12, 'Total')
			# print(Municipal_flow.value)
			# print(Agriculture_flow.value)
			# print(Agriculture2_flow.value)
			# print(Industrial_flow.value)
			# print(PowerPlant_flow.value)
			# print(Indian_flow.value)
			flow_year = {}
			flow_year['Municipal'] = Municipal_flow.value
			flow_year['Agriculture'] = Agriculture_flow.value
			flow_year['Agriculture2'] = Agriculture2_flow.value
			flow_year['Industrial'] = Industrial_flow.value
			flow_year['PowerPlant'] = PowerPlant_flow.value
			flow_year['Indian'] = Indian_flow.value
			flow_senario[str(year)] = flow_year
		flow[scenarios[j]] = flow_senario

	win32com.CoUninitialize()
	sites = ['Municipal', 'Agriculture', 'Agriculture2', 'Industrial', 'PowerPlant', 'Indian']
	source = {'\\from CAPWithdral': 'CAP to', '\\from GW': 'GW to', '\\from SRPwithdral': 'SRP to',
	          '\\from WWTP': 'WWTP to'}
#######################################################################################################
	"""
	Sort the flow value in the following data structure
		{
			"scenario":{
				"name": "",
				"value": [],
				"delta_to_reference": [],
				"site": "",
				"source": ""
			}
			
		}
	"""

	value = {}
	for s in scenarios:
		value_year = []
		for site in sites:
			for l in link:
				var = []
				for y in range(start_year, end_year):
					var.append(flow[s][str(y)][site][l])
				# print(flow[s][str(y)][site][l])
				# print(var)
				value_year.append(
					{'name': source[l] + ' ' + site, 'value': var, 'site': site, 'source': source[l],
					 'format': 'series'})
		value[s] = value_year
		value['timeRange'] = [start_year, end_year]
	# Add percentage change compared to "Reference" scenario
	for s in scenarios:
		for i in range(len(value[s])):
			v = []
			for j in range(len(value[s][i]["value"])):
				v.append("{:.1%}".format((value[s][i]["value"][j] - value[scenarios[0]][i]["value"][j]) / (
							value[scenarios[0]][i]["value"][j]+0.1)))
			value[s][i]["delta_to_reference"] = v
	# print(value[scenarios[2]][2]["delta_to_reference"])
	return value


get_WEAP_flow_value()

# value = WEAP.Branch('\Demand Sites\Municipal').Variables('Annual Activity Level').Expression
# # print(type(value))
#
# with open('data.json', 'w') as file:
# 	json.dump(flow, file)
# file.close()
