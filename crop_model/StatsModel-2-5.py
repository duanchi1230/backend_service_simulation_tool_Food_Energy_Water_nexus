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
def projection(data_test,data_pred,Rpath,writeOut,out):
    data=data_test
    try:
        Year=data_pred['Year']
        varlist=['prop_cotton','prop_alfalfa','prop_corn','prop_barley','prop_durum','prop_veg','prop_remaining']
        pred_matrix_p = pd.DataFrame(np.zeros((data_test.shape[0],len(varlist))),columns=varlist)
        pred_matrix_r = pd.DataFrame(np.zeros((data_test.shape[0],len(varlist))),columns=varlist)

        pred_matrix_p_pr = pd.DataFrame(np.zeros((data_pred.shape[0],len(varlist))),columns=varlist)
        pred_matrix_r_pr = pd.DataFrame(np.zeros((data_pred.shape[0],len(varlist))),columns=varlist)
        #print(data)
        #r = pr.R(RCMD="C:/Program Files/R/R-3.6.3/bin/R")
        #r = pr.R(RCMD="C:/Program Files/R/R-4.0.2/bin/R")
        r = pr.R(RCMD=Rpath)

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
            #print(r("summary(mod)"))
            #
            #print("Prediction ","for",str(i),"by using R \n")
            #print(r('predict(mod,newdata=r_data_test, type="response")[1:(nrow(data)/2)]'))

            r('a=predict(mod,newdata=r_data_test, type="response")[1:nrow(r_data_test)]')
            #r('write.csv(a,"C:/Users/Chiranjib/Desktop/Assignment/Mayuri/output_temp.csv",row.names=FALSE)')
            r('write.csv(a,"C:/Data/R/out_temp.csv",row.names=FALSE)')
            #output = pd.read_csv('C:/Users/Chiranjib/Desktop/Assignment/Mayuri/output_temp.csv')
            output = pd.read_csv('C:/Data/R/out_temp.csv')
            pred_matrix_r[[i]]=output[['x']]
            r('a=predict(mod,newdata=r_data_pred, type="response")[1:nrow(r_data_pred)]')
            r('write.csv(a,"C:/Data/R/out_temp.csv",row.names=FALSE)')
            #output = pd.read_csv('C:/Users/Chiranjib/Desktop/Assignment/Mayuri/output_temp.csv')
            output = pd.read_csv('C:/Data/R/out_temp.csv')

            pred_matrix_r_pr[[i]]=output[['x']]
            #

            #
        a=pred_matrix_r_pr.sum(axis = 1)  # all sums are almost 1
        a=a.tolist()
        a=np.asarray(a)
        a=a.reshape((60,1))
        p1=pred_matrix_r_pr['prop_cotton'].tolist()
        p2=pred_matrix_r_pr['prop_alfalfa'].tolist()
        p3=pred_matrix_r_pr['prop_corn'].tolist()
        p4=pred_matrix_r_pr['prop_barley'].tolist()
        p5=pred_matrix_r_pr['prop_durum'].tolist()
        p6=pred_matrix_r_pr['prop_veg'].tolist()
        p7=pred_matrix_r_pr['prop_remaining'].tolist()
        allval=np.empty((60,7))
        for i in range(60):
            for j in range(7):
                if j==0:
                    allval[i,j]=np.divide(p1[i],a[i,0])
                elif j==1:
                    allval[i,j]=np.divide(p2[i],a[i,0])
                elif j==2:
                    allval[i,j]=np.divide(p3[i],a[i,0])
                elif j==3:
                    allval[i,j]=np.divide(p4[i],a[i,0])
                elif j==4:
                    allval[i,j]=np.divide(p5[i],a[i,0])
                elif j==5:
                    allval[i,j]=np.divide(p6[i],a[i,0])

                else:
                    allval[i,j]=np.divide(p7[i],a[i,0])

        df = pd.DataFrame(allval)
        #df=df.append(allval)
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
def Main(pathT,pathP,Rpath,writeOut,out):
 #---------------------
 data_test=readTestCSV(pathT)
 data_pred=readProjCSV(pathP)
 #
 try:
  projection(data_test,data_pred,Rpath,writeOut,out)
  print('Close file: no errors-program complete')
 except:
  print('Errors- Program did not finish')
#================================================
#
#
# ============================================================
"""
 Call Main
 last write 01.29.20,06.03.20,07.21.20,07.23.20
 Define the data file to use and output file-
"""
#pathTfile='C:/Users/Chiranjib/Desktop/Assignment/Mayuri/Averagefinal.csv'
#pathPfile='C:/Users/Chiranjib/Desktop/Assignment/Mayuri/FuturedataN.csv'
#out='C:/Users/Chiranjib/Desktop/Assignment/Mayuri/outPuts.csv'
"""
  Where your input data reside next three lines), and where you want to write outputs. If you
  don't want to write an output file, set the boolean variable writeOut (below)
  to FALSE
"""
pathTfile='../MyScripts/MPM/Sampson/21July2020/hold/Averagefinal.csv'
pathPfile='../MyScripts/MPM/Sampson/21July2020/FuturedataN.csv'
out='../MyScripts/MPM/Sampson/21July2020/outPuts.csv'
#
"""
  This (below) is the path to your R program and version number
"""
Rpath='C:/Program Files/R/R-4.0.2/bin/R'
"""
 Chi;
  YOU WILL NEED to hard code a temporary directory and file name for
  a temporary csv file that the R code uses. I could not make it
  generic; R did not like it. lines 111, 113, 116, and 118
"""
#tempPath='C:/data/R/out_temp.csv'
# Write an output file of the results - True or False
writeOut=True
# Call the program using Main(var1,var2)
#Main(pathTfile,pathPfile,writeOut,out)
Main(pathTfile,pathPfile,Rpath,writeOut,out)
# ============================================================
# E.O.F.
