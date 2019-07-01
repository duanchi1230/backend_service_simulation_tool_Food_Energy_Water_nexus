import win32com.client

### This module is still under development###
LEAP = win32com.client.Dispatch('LEAP.LEAPApplication')
LEAP.ActiveArea = LEAP.Areas('Phoenix AMA')
# LEAP.ResultValue()
LEAP.Branch(1)
print(
	LEAP.Branch('Demand\Water\Treatment and Distribution\Surface Water\CAP\White Tanks WTP').Variable('Activity Level'))
