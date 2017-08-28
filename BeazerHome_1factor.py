# -*- coding: utf-8 -*-
"""
CDS Research
Author: Jordan Giebas
Advisor: Dr. Albert Cohen
Michigan State University
Feb. 19, 2017
"""

# Working with Excel
import xlrd

# Datetime Module for time-series data
import datetime

# Mathematical Computations
import scipy
import numpy as np
import math
from scipy.stats import norm
from scipy.optimize import fsolve, root
#from scipy.optimize import curve_fit


def averageOfList( List1 ):
        
    return sum(List1)/float(len(List1))

# Input Param: 
def dateToQuarter( dateString ):

    Q1_list = ["01", "02", "03"]    
    Q2_list = ["04", "05", "06"] 
    Q3_list = ["07", "08", "09"]    
    
    L = dateString.split("-")
    Month = L[1]
    #print("Month: ", Month)
    Year = L[2]
    #print("Year: ", Year)
    
    if Month in Q1_list:
        
        return "1." + Year
    
    elif Month in Q2_list:
        
        return "2." + Year

    elif Month in Q3_list:
        
        return "3." + Year
        
    else:
    
        return "4." + Year       
        
## E^Market  = A*phi(d1) - N*e^M*phi(d2)
def fsolve_function( init, E_market, sigma_E_market, N, r, M, A ):
    
    #A = init[0]
    sigma_A = init[0]

    #Aux functions
    d1 = (math.log(A/N) + (r + 0.5*(sigma_A**2)))/(sigma_A*math.sqrt(M))
    d2 = (math.log(A/N) + (r - 0.5*(sigma_A**2)))/(sigma_A*math.sqrt(M))
    
    #return
    out = [A*norm.cdf(d1) - N*math.exp(-1*(r*M))*norm.cdf(d2) - E_market]
    #out.append(sigma_A*A*norm.cdf(d1) - sigma_E_market*E_market)

    return out


def delta_spread( calib_A, calib_sigma_A, E_market, sigma_E_market, N, r, M, spd ):
    
    d1 = (math.log(calib_A/N) + (r + 0.5*(calib_sigma_A**2)))/(calib_sigma_A*math.sqrt(M))
    d2 = (math.log(calib_A/N) + (r - 0.5*(calib_sigma_A**2)))/(calib_sigma_A*math.sqrt(M))
    
    """
    print("\n======DEBUG=======\n")
    print("calib_A: ", calib_A)
    print("calib_sigma_A: ", calib_sigma_A)
    print("E_market: ", E_market)
    print("sigm_E_mkt: ", sigma_E_market)
    print("N: " , N)
    print("r: ", r)
    print("spd: ", spd)  
    """
    
    
    spread_calib2 = (-1.0/M)*math.log( norm.cdf(d2) + ((calib_A*math.exp(r*M))/N)*norm.cdf(-1.0*d1) )    
    
    #print("calculated spread, is it %age?: ", spread_calib2)    
    
    d_spread = (spd-spread_calib2)
    
    #return (spd-spread_calib2)
    return [ d_spread, spread_calib2 ]
    
    

#establish the map between quarter and paramteres for Beazer Home
# map[quarter] = list[ STD, LTD, B (liabilities), E (Equity), sigma_E (volatility), r (10 Yr Treasury Rate), Credit_Spread ]
quarterToData = dict()

# Open the Excel Workbook, read in short term debt
book = xlrd.open_workbook( "BeazerHome_AllDataFile.xlsx" )
STD_data = book.sheet_by_index(0) # sheet containing Short Term Debt Data
LTD_data = book.sheet_by_index(1) # sheet containing Long Term Debt Data
Lib_data = book.sheet_by_index(2) # sheet containing Liability Data
Eqt_data = book.sheet_by_index(3) # sheet containing Equity Data
Vol_data = book.sheet_by_index(4) # sheet containing Volatility Data
Rte_data = book.sheet_by_index(5) # sheet containing 10 Yr Treasury Data
Spd_data = book.sheet_by_index(6) # sheet containing Spread Data
Bta_data = book.sheet_by_index(7) # sheet containing BetaEM Data

quarters = STD_data.col_slice( colx=1, start_rowx=1 )
nEntries = len(quarters)

Quarters = list()

for i in range (0, nEntries+1):
    
    quarter_i = str( STD_data.cell_value( i, 0 ) )
    quarter_i_formatted = quarter_i[2] + "." + quarter_i[4:]
    
    Quarters.append( quarter_i_formatted )
    
    # map all the quarters to a list containing STD
    quarterToData[ quarter_i_formatted ] = [ float( STD_data.cell_value( i, 1 ) ) ]
    
    # append the LTD data
    quarterToData[ quarter_i_formatted ].append( float( LTD_data.cell_value( i, 1 ) ) )

    # append the Liability data
    quarterToData[ quarter_i_formatted ].append( float( Lib_data.cell_value( i, 1 ) ) )
        
    # append the Equity data
    quarterToData[ quarter_i_formatted ].append( float( Eqt_data.cell_value( i, 1 ) ) )
  
    
#############################################
# Now we must handle the volatility
# and interest rate data, which is daily.
# We take an average over the quarter to 
# proxy these values for the quarter
#############################################

# initialize the two maps we'll need to 
# get corresponding averages
q_vol = dict()
q_rte = dict()
q_spd = dict()
q_bta = dict()

# process the volatlity/treasury rates
dates_Vol = Vol_data.col_slice( colx=0, start_rowx=0 )
nEntries_Vol = len(dates_Vol)

dates_Rte = Rte_data.col_slice( colx=0, start_rowx=0 )
nEntries_Rte = len(dates_Rte)

dates_Spd = Spd_data.col_slice( colx=0, start_rowx=0 )
nEntries_Spd = len(dates_Spd)

dates_Bta = Bta_data.col_slice( colx=0, start_rowx=0 )
nEntries_Bta = len(dates_Bta)

# set up the q_vol map
for i in range (0, nEntries_Vol):
    
    # get the date, however this is a float
    date_i_float  = Vol_data.cell_value( i, 0 )
    
    # convert the floating point date to a datetime object
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) )
    
    # convert the datetime ojbect to a string, and return quarter 
    # from the dateToQuarter function
    quarter_i = dateToQuarter( date_i.strftime('%d-%m-%Y') )

    try:
        
        q_vol[ quarter_i ].append( float( Vol_data.cell_value(i,1) ) )
    
    except KeyError:
    
        # put in the volatility values to the map
        q_vol[ quarter_i ] = [ float( Vol_data.cell_value(i,1) ) ]
        
# get the average volatility over each quarter
for k in q_vol.keys():

    avg = averageOfList( q_vol[k] )
    q_vol[k] = avg
    
#print("\n========= Q_VOL MAP =========\n")
#print(q_vol)


# set up the q_rte map
for i in range (0, nEntries_Rte):
    
    # get the date, however this is a float
    date_i_float  = Rte_data.cell_value( i, 0 )
    
    # convert the floating point date to a datetime object
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) )
    
    # convert the datetime ojbect to a string, and return quarter 
    # from the dateToQuarter function
    quarter_i = dateToQuarter( date_i.strftime('%d-%m-%Y') )
    #print("Quarter: ", quarter_i)
        
    try:
        
        q_rte[ quarter_i ].append( float( Rte_data.cell_value(i,1) ) )
    
    except KeyError:
    
        # put in the volatility values to the map
        q_rte[ quarter_i ] = [ float( Rte_data.cell_value(i,1) ) ]

# get the average volatility over each quarter
for k in q_rte.keys():
    
    avg = averageOfList( q_rte[k] )
    q_rte[k] = avg

#print("\n========= Q_RTE MAP =========\n")
#print(q_rte)


# set up the q_spd map
for i in range (0, nEntries_Spd):
    
    # get the date, however this is a float
    date_i_float  = Spd_data.cell_value( i, 0 )
    
    # convert the floating point date to a datetime object
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) )
    
    # convert the datetime ojbect to a string, and return quarter 
    # from the dateToQuarter function
    quarter_i = dateToQuarter( date_i.strftime('%d-%m-%Y') )
    #print("Quarter: ", quarter_i)
        
    try:
        
        q_spd[ quarter_i ].append( float( Spd_data.cell_value(i,1) ) )
    
    except KeyError:
    
        # put in the volatility values to the map
        q_spd[ quarter_i ] = [ float( Spd_data.cell_value(i,1) ) ]

# get the average volatility over each quarter
for k in q_spd.keys():
    
    avg = averageOfList( q_spd[k] )
    q_spd[k] = avg

#print("\n========= Q_SPD MAP =========\n")
#print(q_spd)

# set up the q_spd map
for i in range (0, nEntries_Bta):
    
    # get the date, however this is a float
    date_i_float  = Bta_data.cell_value( i, 0 )
    
    # convert the floating point date to a datetime object
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) )
    
    # convert the datetime ojbect to a string, and return quarter 
    # from the dateToQuarter function
    quarter_i = dateToQuarter( date_i.strftime('%d-%m-%Y') )
    #print("Quarter: ", quarter_i)
        
    try:
        
        q_bta[ quarter_i ].append( float( Bta_data.cell_value(i,1) ) )
    
    except KeyError:
    
        # put in the volatility values to the map
        q_bta[ quarter_i ] = [ float( Bta_data.cell_value(i,1) ) ]

# get the average volatility over each quarter
for k in q_bta.keys():
    
    avg = averageOfList( q_bta[k] )
    q_bta[k] = avg

#print("\n========= Q_Beta MAP =========\n")
#print(q_bta)

# put the volatility, 10yr rate, and credit spread in the data map
for k in quarterToData.keys():

    try:
        quarterToData[k].append( q_vol[k] )
        quarterToData[k].append( q_rte[k] )
        quarterToData[k].append( q_spd[k] )
        quarterToData[k].append( q_bta[k] )
        
    except KeyError:
        
        continue

#print("\n========= Q_DATA MAP =========\n")
#print(quarterToData)


del quarterToData["4.2015"]

#############################################
# Now that all the data is centralized,
# use the paramters to do the computations
#############################################

#A_init = 70
sigmaA_init = 0.2
init_guess = [sigmaA_init]
M = 5.0 # assume

# get x/y coordinates from computation
# will use these lists for regression of c0, c1
X_coord = []
Y_coord = []
pos_neg = []


for k in quarterToData.keys():
        
    #define input parameters for calibration    
    #multiplied by 10^6 for units
    STD = quarterToData[k][0]*1000000
    LTD = quarterToData[k][1]*1000000
    B = quarterToData[k][2] * 100000
    E = quarterToData[k][3] * 1000000
    sig_E = quarterToData[k][4]
    r = quarterToData[k][5]/100.0
    spd = quarterToData[k][6]*100
    bta = quarterToData[k][7]
    
    
    #Assets = E + B
    A = B + E    
    
    #Notional Calulation
    N = STD + 0.5*LTD

    try:
        
        CalibSigA = fsolve( fsolve_function, init_guess, args=(E, sig_E, N, r, M, A) )
        
        #Aux
        d1 = (math.log(A/N) + (r + 0.5*(CalibSigA**2)))/(CalibSigA*math.sqrt(M))
        d2 = (math.log(A/N) + (r - 0.5*(CalibSigA**2)))/(CalibSigA*math.sqrt(M))        
        
        #print("\n===========Calibration===========\n")
        #print("A (calibrated): ", CalibA)
        #print("sigA (calibrated): ", CalibSigA)   
        #print("A (B+E): ", A)
        d_spread = float( delta_spread( A, CalibSigA, E, sig_E, N, r, M, spd )[0] )
        spread_calib = float( delta_spread( A, CalibSigA, E, sig_E, N, r, M, spd )[1] )
    
        
        
        #pos_neg.append( spread_calib )        
        #print("Delta Spread: ", d_spread)        
        x = float(( (norm.cdf(-d1)*E*bta) / (B*float(norm.cdf(d1))) ))     
        
        if bta >= 1.0 and bta <= 2.0:
            X_coord.append( bta )        
            Y_coord.append( d_spread )
        
        
    except ValueError:
        
        continue


print("\n==========\n") #print space before coordinates
#outFile = open('jordan.csv', 'w')       
#print("============ x/y coordinates ============")
for i in range (0, len(X_coord)):
    
    x_str = str(X_coord[i])
    y_str = str(Y_coord[i])
    outString = x_str + "," + y_str
    print( outString )
    #outFile.write(outString)
    
#outFile.close()


"""
pos_cnt = 0
neg_cnt = 0

for elm in pos_neg:
    
    if elm >= 0:
        
        pos_cnt+=1
    
    else:
        
        neg_cnt+=1
"""







