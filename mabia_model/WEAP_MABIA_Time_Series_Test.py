from mabia_model import script_m
import win32com.client
"""
	Module Name: WEAP MABIA Time Series Setting Test
	Purpose: This module is used to test the time series setting through API for WEAP MABIA module
	Status: Finished
"""

# Initialize the dummmy data from script_m.py which is used for MABIA input
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
MaxU = script_m.MaxU
path_1 = '..\mabia_model\GrossMarginID.txt'
path_2 = '..\mabia_model\ID_3.csv'
district, crop, pct_Area, totArea = MaxU(path_1, path_2)
year = [2001, 2005]

# Format the time series data into string which MABIA module could read as input
crop_time_series = {}
for c in crop:
	crop_time_series[c] = ''
crop_time_series[''] = ''
for y in range(2001, 2006):
	for i in range(len(crop)):
		crop_time_series[crop[i]] += str(y)+','+str(pct_Area[i])+','
	crop_time_series[''] += str(y)+','+str(totArea)+','


def set_mabia_timeSeries(crop_time_series):
	"""
	This module set the time series input for WEAP MABIA module
	:param crop_time_series: The time series of all crops for MABIA
	:return: None
	"""
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	scenarios = ['Current Accounts', 'Linkage']
	for s in scenarios:
		WEAP.ActiveScenario = s
		for k in crop_time_series:
			WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\" + k + ": Area").Expression = \
			'Step('+crop_time_series[k][0:-1]+')'
	win32com.CoUninitialize()



def set_mabia_default():
	"""
	This module set the WEAP MABIA module to default value
	:return:
	"""
	win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	scenarios = ['Current Accounts', 'Linkage']
	for s in scenarios:
		WEAP.ActiveScenario = s
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
			WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\" + s + ": Area").Expression = \
			default_percentage[s]
		WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\: Area").Expression = default_total_area
	win32com.CoUninitialize()
# Set MABIA to default
set_mabia_default()

# Set MABIA to time series
# set_mabia_timeSeries(crop_time_series)