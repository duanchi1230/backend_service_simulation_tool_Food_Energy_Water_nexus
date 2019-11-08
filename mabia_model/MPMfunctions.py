# Created on 10.09.19
# author: David Arthur Sampson
# Updated on: 10.10.19,10.11.19,10.16.19,10.17.19

# This code uses crop yield data and meteorological data to estimate the crop area for selected
# crops in Maricopa County, AZ. The "model" will be used by the InFEWS project at ASU, with
# PI's Dr's White, Mascaro, Aggarwal, Maciejewski (no particular order) (& Dr. Hessam)

# Logistic Regression is a Machine Learning classification algorithm that is used to predict the probability of a categorical dependent variable.
# The independent variables are linearly related to the log odds.

# statsmodels.api source can be obtained from: git clone git://github.com/statsmodels/statsmodels.git. NOTE: "git pull" will update
# I installed: statsmodels-0.10.1-cp37-none-win32.whl
# USE: python -m pip install awesome_package to install a whl or other package using PIP

import pandas as pd
import statsmodels.api as sm
import numpy as np
import csv
import win32com.client

# Temporary data file for pdsi, temperature, precipitation, and yield data
# y_crop is Yield; For cotton: lb/acre. For alfalfa: Tons/acre. For corn, barley, spring wheat and winter wheat: bushels/acre.
# temp = mean annual temperature in degrees F. The mean was taken for past 3 years-
# precipitation = mean annual precipitaiton in inches per year - past three
# pdsi = palmer drought severity index
# P_crop is price. "The units for each crop are same as noted above for yields: from Rimjhim-- how can price be a unit of volume (DAS question)? 
# ===============================================================================================================================================
def writeCSV(FID):
	result = False
	A = np.array(cotton)
	B = np.array(corn)
	C = np.array(barley)
	D = np.array(wwheat)
	E = np.array(alfalfa)
	F = np.array(remaining)
	#
	x = len(A)
	if 0 < x:
		output = np.column_stack((A.flatten(), B.flatten(), C.flatten(), D.flatten(), E.flatten(), F.flatten()))
	try:
		np.savetxt(FID, output, delimiter=',')
		result = True
		return result
	except:
		print("Error in CSV export function")


# =====================================================

def Logit():
	# pathFID='..\MyScripts\MPM\Average3.csv'
	# data = pd.read_csv(pathFID)
	#
	# Arrays for use in the coupled model- individual crops
	global cotton, corn, barley, durum, wwheat, alfalfa, remaining
	cotton = []
	corn = []
	barley = []
	durum = []
	wwheat = []
	alfalfa = []
	remaining = []
	# -----------------------
	# Seven crops examined
	varlist = ['cotton', 'corn', 'barley', 'durum', 'wwheat', 'alfalfa', 'remaining']
	# Loop over crops in hte variable list
	# =====================================
	try:
		for i in varlist:
			Y = data[[i]]
			X = data[
				['y_cotton', 'y_corn', 'y_barley', 'y_durum', 'y_wwheat', 'y_alfalfa', 'p_cotton', 'p_corn', 'p_barley',
				 'p_durum', 'p_wwheat', 'p_alfalfa', 'pdsi', 'temp', 'precipitation']]
			#
			X = np.asarray(X)
			# Code for fm logit
			# class statsmodels.discrete.discrete_model.Logit(endog, exog, **kwargs)[source]
			# ------------------
			mod = sm.Logit(Y, X)
			#

			# Fit the model using maximum likelihood
			# -----------
			result = mod.fit()
			# print(result.summary())
			# print(result.params)

			# Logit Marginal Effects
			# -------------------------
			margeff = result.get_margeff()
			# print(margeff.summary())

			ypred = result.predict(X)
			# print(i,"I")
			# print("Prediction ","for --",str(i),"\n")
			# print(ypred, " each year")
			# print("--------------- STOP ------------------")

			# - messy code. Can I improve this?
			if i == "cotton":
				cotton = list(ypred)
			elif i == 'corn':
				corn = list(ypred)
			elif i == 'barley':
				barley = list(ypred)
			elif i == 'durum':
				durum = list(ypred)
			elif i == 'wwheat':
				wwheat = list(ypred)
			elif i == 'alfalfa':
				alfalfa = list(ypred)
				# print(alfalfa, "what the heck is this?")

			elif i == 'remaining':
				remaining = list(ypred)
				#
	except:
		print("Error in the Logit function call")


def readCSV(pathFID):
	global data
	try:
		data = pd.read_csv(pathFID)
	except:
		print("Script failed to read in the CSV file")


def arrays():
	print('Array function')


# ===============================================
def Main(path, writeOut):
	readCSV(path)
	Logit()
	arrays()
	if writeOut:
		outFID = 'out.csv'
		writeCSV(outFID)

def set_crop_area(start_year, end_year):
	# ================================================
	#
	# Define the data file to use - three choices
	path = 'DataTo2050median.csv'
	# Write an output file of the results - True or False
	writeOut = False
	# Call the program using Main(var1,var2)
	Main(path, writeOut)

	varlist = ['cotton', 'corn', 'barley', 'durum', 'wwheat', 'alfalfa', 'remaining']
	area = [cotton, corn, barley, durum, wwheat, alfalfa, remaining]
	# win32com.CoInitialize()
	WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
	scenarios = ['Current Accounts', 'Linkage']
	for i in range(len(varlist)):
		time_series = ''
		for y in range(start_year, end_year+1):
			time_series = time_series + str(y) +','+ str(area[i][y-1991]*100) + ','
		time_series = 'Step(' + time_series[0:-1] +')'
		print(time_series)
		for s in scenarios:
			WEAP.ActiveScenario = s
			WEAP.BranchVariable("\Demand Sites and Catchments\Agricultural Catchment\\" + varlist[i] + ": Area").Expression = time_series
	win32com.CoUninitialize()


set_crop_area(2001, 2005)

