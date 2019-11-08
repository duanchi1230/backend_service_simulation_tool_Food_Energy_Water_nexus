# Date: Last write was 10.17.19
# Dr. David Arthur Sampson: david.a.sampson@asu.edu
#
# This code was adapted from python script written by Mayuri Roy Choudhury, a student of Dr. Rimjhim Aggarwal
#
# The functions in this script read a csv file of crop data, with the current fields denoted as:
# Year	y_cotton y_corn	y_barley y_durum y_wwheat y_alfalfa p_cotton p_corn p_barley p_durum p_wwheat p_alfalfa	pdsi 
# temp precipitation cotton alfalfa corn barley durum wwheat remaining
#
# NOTE: the prefix for the individual crops "y" is yield: Units are: as reported in USDA data. 
# For cotton: lb/acre; alfalfa: Tons/acre; corn, barley, spring wheat and winter wheat: bushels/acre.
# The prefix "p" is price: "The units for each crop are same as noted above for yields." From Dr. Aggarwal. THIS I do NOT understand David Arthur
#
# temp (temperature) is mean annual temperature in degrees F.; Precipitation: mean annual precipitation. These data came from
# Ballinger, A., Kunkel, K. 2019. Scenarios of Climate Extremes: Phoenix, AZ. Urban Resilience to Extreme Events (internal report). A copy of which
# was placed in the DropBox folder: Sampson/MPMcode/ 
#
# PDSI is the palmer drought severity Index. Sampson created these data as:
# *https://www.ncdc.noaa.gov/cag/statewide/time-series/2/pdsi/all/5/1919-2019
# NOTE: I (Sampson) took the mean from 1920 to 1934 and subtracted the mean from 2000 to 2014 to obtain the difference. I then
# subtracted that difference from the 1920 to 2019 data record and appended these data to the 2000 to 2060 T and precip data
#
# The fields "cotton, alfalfa, corn, barley, durum, wwheat, and remaining" represent the estimate the proportion of total crop area for Maricopa County
# in each of the six plus one crop(s) examined. The "remaining" crop is supposed to represent fruits and vegetables.
#
# The years 1991 to 2018 were used in the development of the multinomial Logit model. Accordingly, all values in the file for that period were
# created by Dr. Aggarwal and her students. Except, I have superimposed my climate data estimates over hers starting in the year 2000
#
# For values 2019 and beyond there were three approaches to "estimate" these data (as of 10.17.19):
# Mean yield and mean prices and mean response from cotton, alfalfa, corn, etc.(over the 1991 to 2018 data set)
#
# Median yield and median prices and median response from cotton, alfalfa, corn, etc. (over the 1991 to 2018 data set)
#
# Median yield and median prices and then the last response (i.e., 2018) from cotton, alfalfa, corn, etc. (over the 1991 to 2018 data set)
#
# The seven arrays contain the estimated proportional area for the crops: "cotton, alfalfa, corn, barley, durum, wwheat, and remaining"
#