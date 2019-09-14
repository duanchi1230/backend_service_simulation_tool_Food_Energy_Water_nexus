import win32com.client

'''
	Module Name: WEAP and LEAP coupling test
	Purpose: This module is used to test the external link through script between WEAP and LEAP
	Status: Still under development
'''

LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')


# print(LEAP.View)
# print(LEAP.Branch(1).name)


def CAP_pumping(year):
	v1 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Power2\\from Withdrawal Node 3:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	v2 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 1:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	return v1 + v2 * 1.5


def WTP(year):
	v1 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 2:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	v2 = WEAP.ResultValue(
		'Supply and Resources\Transmission Links\\to Municipal\\from Withdrawal Node 1:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	return (v1 + v2) * 0.45


def WWTP(year):
	v1 = WEAP.ResultValue(
		'Supply and Resources\Return Flows\\from WWTP\\to WWTP Return:Total Node Outflow', year, 1,
		'Linkage', year,
		12)
	return v1 * 0.53


def Power2(year):
	v1 = LEAP.Branch('Transformation\Electricity generation\Processes\Power2').Variable(
		'Average Power Dispatched').Value(year)
	obj = LEAP.Branch('Transformation\\Electricity generation\\Processes\\Power1').Variable(
		'Average Power Dispatched').Value(year)
	object_methods = [method_name for method_name in dir(obj)
	                  if callable(getattr(obj, method_name))]
	print(obj)
	return v1


print(LEAP.Branch('Transformation\\Electricity generation\\Processes\\Power1').Variable(
		'Exogenous Capacity'))
LEAP.Branch('Transformation\\Electricity generation\\Processes\\Power1').Variable(
		'Exogenous Capacity').Expression = 300

print(WEAP.TimeStepName(1))


LEAP.BaseYear=2002
print(LEAP.BaseYear)
print(LEAP.EndYear)
Power2(2001)
