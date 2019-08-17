import win32com.client
from mabia_model.script_m import MaxU
from mabia_model.testing import hello

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
				path.append({'demand': str(node), 'source': str(name), 'path': (str(node)+'\\'+str(name)) })
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
	### Uninitialize the win32com object ###
	win32com.CoUninitialize()
	return flow, timeRange

def set_mabia_input():
	### Initialize win32com object###
	win32com.CoInitialize()
	### Initialize WEAP application communication port###
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	### Import the MABIA outputs ###
	path_1 = '..\mabia_model\GrossMarginID.txt'
	path_2 = '..\mabia_model\ID_3.csv'
	district, crop, pct_Area, totArea = MaxU(path_1, path_2)
	# print("output" , crop)

	### Set MABIA parameters ###
	for i in range(len(crop)):
		WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\" + crop[i] +": Area").Expression = pct_Area[i]
	WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\: Area").Expression =totArea
	### Uninitialize the win32com object ###
	win32com.CoUninitialize()

### This function works to set the MABIA parameters to default ###
def set_mabia_default():
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	default_percentage = {
		"alfalfa": 50.423,
		"barley": 13.488,
		"sorghum": 1.857,
		"cotton": 21.210,
		"winter_wheat": 0.549,
		"potatoes": 0.132,
		"sugarbeet": 0.439,
		"corn": 8.52,
		"durham_wheat": 3.382,
		"other": 0
	}
	default_total_area = 40429.48
	for s in default_percentage:
		WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\" + s +": Area").Expression = default_percentage[s]
	WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\: Area").Expression =default_total_area
	win32com.CoUninitialize()

# set_mabia_default()
get_WEAP_flow_value()
