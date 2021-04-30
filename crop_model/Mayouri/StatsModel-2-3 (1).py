"""
# Created on 06.03.20
# author: Mayuri
# Updated on: 06.04.20


#Logistic Regression is a Machine Learning classification algorithm that is used to predict the probability of a categorical dependent variable. The independent variables are linearly related to the log odds.

# statsmodels.api source can be obtained from: git clone git://github.com/statsmodels/statsmodels.git. NOTE: "git pull" will update
# I installed: statsmodels-0.10.1-cp37-none-win32.whl
# USE: python -m pip install awesome_package to install a whl or other package using PIP
"""
import win32com.client
import pythoncom
import pandas as pd
from pandas import DataFrame
import statsmodels.api as sm
import numpy as np
import pyper as pr
import csv

"""
Temporary data file for pdsi, temperature, precipitation,

 yield data
 y_crop is Yield; For cotton: lb/acre. For alfalfa: Tons/acre. For corn, barley, spring wheat and winter wheat: bushels/acre.
 temp = mean annual temperature in degrees F. The mean was taken for past 3 years-
 precipitation = mean annual precipitaiton in inches per year - past three
 pdsi = palmer drought severity index
 P_crop is price. "The units for each crop are same as noted above for yields:
 i.e., $ per buschel, $ per ton, $ per pound

 Outputs are the proportional area for each crop
"""

"""
NOTE: this code is a direct representation of python code that I obtained from
 Ag on 06.03.20 transposed from code that Mayuri  wrote.

I modified the code to make it modualized, more easily read and debug, and used
as generic code (transferble) for the coupled inFEWs model

Code was created from finalpy_final.py with the following header-
Created on Fri Oct  4 18:45:22 2019

@author: Mayuri
Updated on June 1 , 12.00pm MST
"""
#====================================================================================

#=======================
# June 2020 Mayuri code
def projection(data_test,data_pred,writeOut,out):
    data=data_test
    try:
        Year=data_pred['Year']
        varlist=['prop_cotton','prop_alfalfa','prop_corn','prop_barley','prop_durum','prop_veg','prop_remaining']
        pred_matrix_p = pd.DataFrame(np.zeros((data_test.shape[0],len(varlist))),columns=varlist)
        pred_matrix_r = pd.DataFrame(np.zeros((data_test.shape[0],len(varlist))),columns=varlist)

        pred_matrix_p_pr = pd.DataFrame(np.zeros((data_pred.shape[0],len(varlist))),columns=varlist)
        pred_matrix_r_pr = pd.DataFrame(np.zeros((data_pred.shape[0],len(varlist))),columns=varlist)

        #print(data)
        r = pr.R(RCMD="C:/Program Files/R/R-4.0.2/bin/R")

        for i in varlist:
            #print(i)
            Y=data[[i]]

            X=data[['y_cotton','y_corn','y_barley','y_durum','y_alfalfa','p_cotton','p_corn','p_barley','p_durum','p_alfalfa','temp','precipitation','pdsi']]
            X=np.asarray(X)
            #Code for fm logit
            mod = sm.Logit(Y, X)
            res=mod.fit(cov_type='HC0')
            Xtest=data_test[['y_cotton','y_corn','y_barley','y_durum','y_alfalfa','p_cotton','p_corn','p_barley','p_durum','p_alfalfa','temp','precipitation','pdsi']]
            Xtest=np.asarray(Xtest)
            ytest=res.predict(Xtest)
            Xpred=data_pred[['y_cotton','y_corn','y_barley','y_durum','y_alfalfa','p_cotton','p_corn','p_barley','p_durum','p_alfalfa','temp','precipitation','pdsi']]
            #Xpred=data_pred[['y_cotton','y_corn','y_barley','y_durum','y_alfalfa','y_vegetables','p_cotton','p_corn','p_barley','p_durum','p_alfalfa','temp','precipitation','pdsi']]


            Xpred=np.asarray(Xpred)
            ypred=res.predict(Xpred)
            """
                Python summary - uncomment the following line
            """
            #print(res.summary())
            #
            #print("Prediction ","for",str(i),"\n")
            #print(ypred)
            pred_matrix_p[[i]]=pd.DataFrame(ytest)
            pred_matrix_p_pr[[i]]=pd.DataFrame(ypred)
            #
            data['w']=Y
            r.r_data = data
            r.r_data_test = data_test
            r.r_data_pred = data_pred
            r('')
            r('data <- rbind(cbind(r_data, y = 1, wt = r_data$w), cbind(r_data, y = 0, wt = 1 - r_data$w))')

            r('mod <- glm(y ~ y_cotton+y_corn+y_barley+y_durum+y_alfalfa+p_cotton+p_corn+p_barley+p_durum+p_alfalfa+temp+precipitation+pdsi, weights = wt, subset = (wt > 0), data = data, family = binomial)')
            print(r("summary(mod)"))
            #
            #print("Prediction ","for",str(i),"by using R \n")
            #print(r('predict(mod,newdata=r_data_test, type="response")[1:(nrow(data)/2)]'))

            r('a=predict(mod,newdata=r_data_test, type="response")[1:nrow(r_data_test)]')
            r('write.csv(a,"C:/Data/R/output_temp.csv",row.names=FALSE)')

            output = pd.read_csv('C:/Data/R/output_temp.csv')

            pred_matrix_r[[i]]=output[['x']]

            r('a=predict(mod,newdata=r_data_pred, type="response")[1:nrow(r_data_pred)]')
            r('write.csv(a,"C:/Data/R/output_temp.csv",row.names=FALSE)')
            output = pd.read_csv('C:/Data/R/output_temp.csv')
            pred_matrix_r_pr[[i]]=output[['x']]
            #
            """
                Add the object summary statment below
            """
            #
            #
            df = pd.DataFrame(pred_matrix_r)
            df=df.append(pred_matrix_r_pr)
            df.insert(0,"Year",Year)
            #df = pd.DataFrame(pred_matrix_p)
            #df=df.append(pred_matrix_p_pr)
            #
            #df.insert(0,"Year",Year)

        if(writeOut):
            try:
                df.to_csv(out,index=False, header=True,sep=',')
                print('writing outputs')
            except:
                print('Error writing the output file')
        else:
          print('No outputs written to file')

    except:
        print("Error in the Projection function")
# =======================================================
#
#========================
def readTestCSV(pathFID):

 try:
  data = pd.read_csv(pathFID)
  return data
 except:
  print("Script failed to read in the test CSV file")
#===================================================
#
#============================
def readProjCSV(pathFID):

 try:
  data = pd.read_csv(pathFID)
  return data
 except:
  print("Script failed to read in the projection CSV file")
#===========================================================
#
#
# ===============================================
def Main(pathT,pathP,writeOut,out):
 #---------------------
 data_test=readTestCSV(pathT)
 data_pred=readProjCSV(pathP)
 #
 try:
  projection(data_test,data_pred,writeOut,out)
  print('Close file: no errors-program complete')
 except:
  print('Errors- Program did not finish')
#================================================
#
#
# ============================================================
"""
 Call Main
 last write 01.29.20,06.03.20
 Define the data file to use and output file-
"""
pathTfile='..\MyScripts\MPM\Sampson\21July2020\hold\Averagefinal.csv'
#pathTfile='Averagefinal.csv'
pathPfile='FuturedataN.csv'
pathPfile='..\MyScripts\MPM\Sampson\21July2020\FuturedataN.csv'
out='outPuts.csv'
out='..\MyScripts\MPM\Sampson\21July2020\outPuts.csv'
# Write an output file of the results - True or False
writeOut=True
# Call the program using Main(var1,var2)
Main(pathTfile,pathPfile,writeOut,out)
# ============================================================
# E.O.F.
