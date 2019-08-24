import win32com.client

### This module is still under development###
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
# LEAP.ActiveArea = LEAP.Areas('Internal_linking_test')
# LEAP.ResultValue()
print(LEAP.View)
print(LEAP.Branch(1).name)


# print(
# LEAP.Branch('Demand\Water unrelated\Per capita demand').Variable('Activity Level'))

# print(
# 	LEAP.Branch('Demand\Water related\CAP pumping').Variable('Energy Demand Final Units').Value(2002))
# print(
# 	LEAP.Branch('Demand\Water related\WTP').Variable('Energy Demand Final Units').Value(2002))
# print(
# 	LEAP.Branch('Transformation\Electricity generation\Processes\Power2').Variable('Average Power Dispatched').Value(2002))

# print(WEAP.BranchVariable("\Demand Sites and Catchments\Power2\\: Monthly Demand").Expression)
#
# print(LEAP.Branch('Demand\Water related\CAP pumping').Variable('Final Energy Intensity Time Sliced').Expression)


def CAP_pumping(year):
	v2 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Power2\\from Withdrawal Node 3:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	v1 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 1:Total Node Outflow', year, 1,
		'Linkage', year,
		12)

	return v1+v2

# (WEAPValue(Supply and Resources\Transmission Links\to Power2\from Withdrawal Node 3:Total Node Outflow[m^3])+
# WEAPValue(Supply and Resources\Transmission Links\to Municipal\from Withdrawal Node 1:Total Node Outflow[m^3]))*1.5

CAP_pumping(2002)
