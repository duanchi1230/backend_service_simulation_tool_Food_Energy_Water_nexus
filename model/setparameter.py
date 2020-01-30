import pythoncom
import win32com.client
WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
start_year = WEAP.BaseYear
end_year = WEAP.EndYear
WEAP.ActiveArea = 'Ag_MABIA_v14'