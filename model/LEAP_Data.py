import win32com.client

### This module is still under development and is used to extract LEAP results###
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
# LEAP.ActiveArea = LEAP.Areas('Internal_linking_test')
# LEAP.ResultValue()
# print(LEAP.View)
# print(LEAP.Branch(1).name)
#
# print(
# 	LEAP.Branch('Demand\Water unrelated\Per capita demand').Variable('Activity Level'))
#
# print(
# 	LEAP.Branch('Demand\Water related\CAP pumping').Variable('Energy Demand Final Units').Value(2002))
# print(
# 	LEAP.Branch('Demand\Water related\WTP').Variable('Energy Demand Final Units').Value(2002))
print(
	LEAP.Branch('Demand\Water related\WWTP').Variable('Energy Demand Final Units').Value(2002))

def get_LEAP_value():
	win32com.CoInitialize()
	LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')

	for b in LEAP.Branches:
		if b.name=='White Tanks WTP':
			print('1', b.Variable('Activity Level'))

get_LEAP_value()