import statistics
import json
import csv
import random
from collections import namedtuple


#
# Python script written by Dr. David arthur Sampson
# ASU Decision Center for a Desert City
# JAW Global Institute of Sustainability
#
# This script serves as a place holder for an Economic 
# model being written by Rimjhim Aggarwal's student(s)
# Used in the inFEWs project headed up by Dr. Maciejewski,
# Dr. White, Dr. Aggarwal, and Dr. Mascaro (& Dr. Hessam)
#
# This code will feed in Crop Area data for the WEAP-MABIA
# model
#
# Last write: 07.12.19,07.23.29
#
# Start
# =========================================================
# Read in data from files for initial conditions
# This includes crop area and gross margin per crop
# ------------------------
def initiate(FID_1, FID_2):
	global crop, values, data
	print(FID_1, FID_2)
	# try:
	if True:
		# read in gross margin per crop and district
		crop, values = extract_grossMargin(FID_1)

		# read in the districts, crops, and  area for each crop
		I2 = extract_new(FID_2)

	# except:
	# 	print("Script failed to initiate input parameters")


# ======================================================

# ======================================================
def MaxU(path1, path2):
	#  risk = [0.9,0.9,0.9,0.9,0.9,0.9,0.9,0.9,0.9,0.9]
	#  i = 0
	#  q = statistics.stdev(z)
	# try:
	if True:
		initiate(path1, path2)
		#
		n = len(data)
		ZeeEst(n)

		input = Zoptimize(n, data)
		Out_2 = writeOutput(data, input)
		Out_3 = pctArea(n)
		Out_4 = areaCheck(n)
		#
		print("MPC MaxU function END")

		# So, input into MABIA would be the district,crop, and cropArea arrays (all indexed the same)
		#  Lists are: district[i],crop[i],pct_Area[i]
		#  total Area of all crops is: totArea
		print(crop)
		return district, crop, pct_Area, totArea
	# except:
	# 	print("Error in the Max U function")


# ======================================================

def ZeeEst(n):
	global Zee
	Zee = []
	temp = []
	std = []
	z = 0
	j = 0
	k = 0
	x = 0
	w = 0
	#

	#
	# This will be the start of an outer loop
	f1 = Z1(n, crop, values)
	#
	sfp = 0
	f2 = Z2(n, data, sfp)
	#
	numDistricts = 1
	f3 = Z3(numDistricts)
	#
	period = 365
	f4 = Z4(period)
	#
	f5 = Z5(period)
	#
	wc = [1000]
	f6 = Z6(numDistricts, wc)
	#
	ia = [1000]
	f7 = Z7(numDistricts, ia)
	#
	risk = [0.9]
	numCrops = 10
	#
	f8 = periodByDistrict(numDistricts)
	#

	m = 19
	n = len(Zee1)
	f9 = calcSum_array(Zee1, n, m)

	gmX = []
	gmX = Sum_array
	f10 = calcSum_array(Zee2, n, m)
	sbX = []
	sbX = Sum_array
	#
	i = 0
	try:
		while i < numDistricts:
			a = float(gmX[i])
			b = float(sbX[i])
			c = float(f8[i])
			z = (a + b) * c
			i += 1

			# s = statistics.stdev(Zee)
			# yhat = y * r * s

	except:
		print("Error in the estimate of Z - general call")


# =====================================================
def JoinArea(n, arr, gross):
	NewArray = []
	for i in arr:
		NewArray[i] = float(arr[i]) / float(gross[i])
		i += 1
	if i > n - 1:
		return NewArray


# =====================================================
def calcSum_array(arr, n, m):
	global Sum_array
	Sum = 0
	Sum_array = [0 for i in range(n)]

	# calc 1st m/2 + 1 element for 1st window
	for i in range(m // 2 + 1):
		Sum += arr[i]
	Sum_array[0] = Sum

	# use sliding window to
	# calculate rest of Sum_array
	for i in range(1, n):
		if (i - (m // 2) - 1 >= 0):
			Sum -= arr[i - (m // 2) - 1]
		if (i + (m / 2) < n):
			Sum += arr[i + (m // 2)]
		Sum_array[i] = Sum

	# -prSum_array
	# for i in range(n):
	#    print(i,Sum_array[i], end = " ")


# =====================================================
def periodByDistrict(num):
	global period
	period = []
	w = 0
	k = 0
	try:
		while k < num:
			w = Zee3[k] * (Zee5[k] - (Zee6[k] - Zee7[k]))
			# w = Zee3[k] * (Zee4[k] - Zee5[k] - (Zee6[k] - Zee7[k]))

			period.append(w)
			k += 1
		if k > num - 1:
			return period
	except:
		print("Error in periodByDistrict")


# ======================================================
# First part of the Z equation
# - Gross Margin and crop area
# ----------------------
def Z1(n, crop, gross):
	global Zee1
	Zee1 = []
	i = 0

	try:
		for area in enumerate(data):
			for G in gross:
				x = float(G)
				a = data[i].Area
				y = float(G) * float(data[i].Area)
				Zee1.append(y)
				i += 1

				if i > n - 1:
					return Zee1

	except:
		print("Error in the Z1 function")


# function that returns a constant value of 1
# ======================================================
def Z2(n, data, sfp):
	global Zee2
	Zee2 = []
	i = 0
	try:
		for area in enumerate(data):
			sb = float(sfp) * float(data[i].Area)
			Zee2.append(sb)
			i += 1

			if i > n - 1:
				Zee2sum = sum(Zee2) + sfp
				return Zee2sum
	except:
		print("Error in the Z2 function")


# function that returns a constant value of 1
# ======================================================
def Z3(n):
	global Zee3
	Zee3 = []
	i = 0
	mdu = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]  # modulation rate
	fco = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  # family labour opportunity cost
	try:
		while i < n:
			x = mdu[i]
			y = fco[i]
			z = float(x) - float(y)
			Zee3.append(z)
			i += 1
			if i > n - 1:
				Zee3sum = sum(Zee3)
				return Zee3sum
	except:
		print("Error in the Z3 function")


# ======================================================
#   
def Z4(n):
	global Zee4
	Zee4 = []
	i = 0
	flab = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  # family labour use per period

	try:
		while i < n:
			for lab in flab:
				x = float(lab)
				z = x
				Zee4.append(z)
				i += 1

				if i > n - 1:
					Zee4sum = sum(Zee4)
					return Zee4sum
	except:
		print("Error in the Z4 function")


# ======================================================
def Z5(n):  # return

	global Zee5
	Zee5 = []
	hlab = [35, 30, 25, 20, 15, 10, 5, 1, 1, 1, 1, 1]  # hired labour per period
	hlw = [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5]  # hired labor wage in $ hour-1
	i = 0
	days = 200
	try:
		while i < n:
			for lab in hlab:
				for hl in hlw:
					y = float(lab) * float(hl) * float(days)
					Zee5.append(y)
					i += 1

					if i > n - 1:
						Zee5sum = sum(Zee5)
						return Zee5sum
	except:
		print("An error in the execution of Z5 occurred")


# ======================================================
def Z6(n, wc):
	global Zee6
	Zee6 = []
	wpm = [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5]  # volumetric water price

	i = 0
	try:
		while i < n:
			for p in wpm:
				y = float(p) * float(wc[i])
				Zee6.append(y)
				i += 1

				if i > n - 1:
					Zee6sum = sum(Zee6)
					return Zee6sum
	except:
		print("An error in the execution of Z6 occurred")


# ======================================================
def Z7(n, ia):
	global Zee7
	Zee7 = []
	fee = [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5]  # irrigation water fee paid per ha

	i = 0
	try:
		while i < n:
			for f in fee:
				y = float(f) * float(ia[i])
				Zee7.append(y)
				i += 1

				if i > n - 1:
					Zee7sum = sum(Zee7)
					return Zee7sum
	except:
		print("An error in the execution of Z7 occurred")


# ======================================================
# ======================================================

# ======================================================
def Zoptimize(n, Data):  # return

	global result
	result = []
	i = 0
	try:
		while i < n:
			x = float(Data[i].Area)
			y = float(random.randint(1, 101))
			z = x
			result.append(z)
			i += 1

			if i > n - 1:
				return result
	except:
		print("An error in the execution of Zoptmize occurred")


# =====================
# Equation variables
# ======================================================
def extract_grossMargin(filename):
	global crop
	global grossMargin
	infile = open(filename, 'r')
	infile.readline()  # skip the first line
	crop = []
	grossMargin = []
	for line in infile:
		words = line.split()
		# words[0]: month, words[1]: rainfall
		crop.append(words[0])
		grossMargin.append(float(words[1]))
	infile.close()
	crop = crop[:]
	grossMargin = grossMargin[:]  # Redefine to contain monthly data
	return crop, grossMargin


#
# =======================================
# Input data for the series of equations
#
# ======================================================
# this works
def extract_Area(filename):
	try:
		input_file = csv.DictReader(open(filename))
		csv_dict = {elem: [] for elem in input_file.fieldnames}
		for row in input_file:
			for key in csv_dict.keys():
				csv_dict[key].append(row[key])
				# print(Tonapah.Crop.alfalfa)
	except:
		print("Error in CSV import")


# ======================================================
# this works too
def extract_newArea(filename):
	global data
	try:
		with open(filename) as f:
			reader = csv.DictReader(f)
			data = [r for r in reader]
	except:
		print("Error in new CSV import")


# ======================================================
# Reads a CSV file with starting values for crop Area
# ------------------------
def extract_new(filename):
	global data
	try:
		with open(filename) as f:
			reader = csv.reader(f)
			top_row = next(reader)
			Data = namedtuple("Data", top_row)
			data = [Data(*r) for r in reader]
	except:
		print("Error in new named tuple CSV import")


# =====================================================


# =====================================================
def writeCSV(Data, input, FID):
	n = 120

	try:
		with open(FID, mode='w', newline='') as csv_file:
			fieldnames = ['District', 'Crop', 'CropArea']
			writer = csv.DictWriter(csv_file, delimiter=',', fieldnames=fieldnames)

			writer.writeheader()
			# while i < n
			writer.writerow({'District': Data[1].District, 'Crop': Data[1].Crop, 'CropArea': input[1]})

			# i+=1

			# if i > n-1:
			# return result

	except:
		print("Error in CSV export function")


# =====================================================
def writeOutput(data, input):
	global district
	global crop
	global cropArea
	district = []
	crop = []
	cropArea = []
	i = 0
	n = len(input)
	yminus1 = 0
	try:
		for area in enumerate(data):
			w = data[i].District
			x = data[i].Crop
			y = input[i]
			district.append(w)
			crop.append(x)
			cropArea.append(y)
			MyMax = max(y, yminus1)
			yminus1 = y
			i += 1

			if i > n - 1:
				# print(MyMax)
				return district, crop, cropArea

	except:
		print("Error in Array as an output for the interface")


# ====================================================
def pctArea(n):
	global pct_Area
	global totArea
	pct_Area = []
	i = 0
	j = 0
	y = 0
	z = 0
	totArea = totalArea(n)

	try:
		while i < n:
			if 0 < totArea:
				y = round((cropArea[i] / totArea) * 100, 1)
				pct_Area.append(y)

				i += 1

	except:
		print("Error in the calculation of the % area for each crop")


# =============================================================

def areaCheck(n):
	i = 0
	j = 0
	w = 0
	z = 0
	try:
		while j < n:
			z = z + pct_Area[j]
			j += 1
			if j == n:
				if z != 100:
					while i < n - 1:
						w = w + pct_Area[i]
						i += 1
						if i == n - 1:
							pct_Area[n - 1] = 100 - w
							s = sum(pct_Area)
							if s != 100:
								print(s, "Rounding error for pct_Area")

	except:
		print("Error in areaCheck")


# ====================================================
def totalArea(n):
	Asum = 0
	try:
		return (sum(cropArea))
	except:
		print("Error in totalArea function call")


# ====================================================
# 
# ====================================================
def testing(n):
	i = 0
	try:
		while i < n:
			print(district[i], crop[i], pct_Area[i], totArea)
			i += 1
			if i == n:
				print("Finished")
	except:
		print("Error in testing function Call")

# ====================================================

# This is where everything is finalized

# path_1 = '..\mabia_model\GrossMarginID.txt'
# path_2 = '..\mabia_model\ID_3.csv'
# out = MaxU(path_1,path_2)

# n = len(data)
# test=testing(n)
#
# ===================================
