import win32com.client
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

	for b in LEAP.Branches:
		LEAP.Branch(b.FullName)
		print('\n')
		print(b.FullName)
		print(b.Name)
		for v in LEAP.Branch(b.FullName).Variables:
			if v != None:
				a = 1
				print('   ', v.name, ':', v.Value(2005))
			# if v.IsResultVariable == True:
			# 	print(v.name)
				# print(type(LEAP.Branch(b.FullName)))
				# path = b.FullName + ":" +v.name
				# print(LEAP.ResultValue(path, 2002, 1, 'Linkage', 2002,12, 'Total'))
	# print(LEAP.Branch("\Demand\Water unrelated\Per capita demand").Variable('Energy Demand Final Units').Value(2002, 'MWH'))
	print(len(LEAP.Branches))
	win32com.CoUninitialize()

def path_parser(path):
	"""
	:param path: A string in the format of 'Transformation\Electricity generation\Output Fuels\Electricity'
	:return: A parsed string array in the format ['Transformation', 'Electricity generation', 'Output Fuels', 'Electricity']
	"""
	branch = []
	name = ''
	for character in path:
		if character !='\\':
			name = name + character
		else:
			if name !='':
				branch.append(name)
				name = ''
	branch.append(name)
	return branch
path = 'Transformation\Electricity generation\Output Fuels\Electricity'
print(path_parser(path))

get_LEAP_Outputs()
