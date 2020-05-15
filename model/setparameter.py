import pythoncom
import win32com.client
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
start_year = WEAP.BaseYear
end_year = WEAP.EndYear
print(WEAP.Timesteps[0])

# LEAP = win32com.client.Dispatch('LEAP.WEAPApplication')
# print(LEAP.CalculationTime)

from datetime import datetime

# datetime object containing current date and time
# now = datetime.now()
# dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
# print("date and time =", dt_string)
# dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
# print("date and time =", dt_string)
# for i in range(10000000):
# 	i+=1
# now = datetime.now()
# dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
# print("date and time =", dt_string)

# run_log_file = open('../run_log_file.txt', 'r')
# print(run_log_file.readlines())
# for line in run_log_file.readlines():
# 	print(line)
