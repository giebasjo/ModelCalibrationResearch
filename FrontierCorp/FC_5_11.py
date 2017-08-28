#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CDS Research
Author: Jordan Giebas
Advisor: Dr. Albert Cohen
Michigan State University
March 16, 2017

Company: Frontier Communications Corp.

"""

import matplotlib.pyplot as plt

# Working with Excel
import xlrd

# Datetime Module for time-series data
import datetime

# Mathematical Computations
import math
from scipy.stats import norm
from scipy.optimize import fsolve


def averageOfList( List1 ):
        
    return sum(List1)/float(len(List1))

def firstDayCheck( day_month_String ):
    
    firstDayList = ["01-01", "01-04", "01-07", "01-10"]
    
    if day_month_String in firstDayList:
        
        return True
    
    return False

def secondDayCheck( day_month_String ):
    
    secondDayList = ["02-01", "02-04", "02-07", "02-10"]
    
    if day_month_String in secondDayList:
        
        return True
    
    return False

def thirdDayCheck( day_month_String ):
    
    thirdDayList = ["03-01", "03-04", "03-07", "03-10"]
    
    if day_month_String in thirdDayList:
        
        return True
    
    return False

def fourthDayCheck( day_month_String ):
    
    fourthDayList = ["04-01", "04-04", "04-07", "04-10"]
    
    if day_month_String in fourthDayList:
        
        return True
    
    return False
    
def notWeekendCheck( dayString ):
    
    weekendList = ["Saturday", "Sunday"]
    
    if dayString in weekendList:
        
        return False
    
    return True

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
def fsolve_function( init, E_market, sigma_E_market, r, M, A):

    sigma_A = init[0]
    N = init[1]

    #Aux functions
    d1 = (math.log(A/N) + (r + 0.5*(sigma_A**2))*M)/(sigma_A*math.sqrt(M))
    d2 = (math.log(A/N) + (r - 0.5*(sigma_A**2))*M)/(sigma_A*math.sqrt(M))
    
    #return
    out = [A*norm.cdf(d1) - N*math.exp(-1*(r*M))*norm.cdf(d2) - E_market]
    out.append(sigma_A*A*norm.cdf(d1) - sigma_E_market*E_market)

    return out


def delta_spread( A, CalibSigA, E_market, sigma_E_market, calibN, r, M, spd ):
    
    d1 = (math.log(A/calibN) + ((r + 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))
    d2 = (math.log(A/calibN) + ((r - 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))    
    
    spread_calib = (-1.0/M)*math.log( norm.cdf(d2) + ((A*math.exp(r*M))/calibN)*norm.cdf(-1.0*d1) )
    spread_calib *= 10000.0
    
    #return (spd-spread_calib2)
    return [(spd-spread_calib), spread_calib]


#establish the map between quarter and paramteres for Beazer Home
# map[quarter] = list[ STD, LTD, B (liabilities), E (Equity), sigma_E (volatility), r (10 Yr Treasury Rate), Credit_Spread ]
quarterToData = dict()

# Open the Excel Workbook, read in short term debt
book = xlrd.open_workbook( "FrontierCorp_1factor.xlsx" )
STD_data = book.sheet_by_index(0) # sheet containing Short Term Debt Data (Quarterly)
LTD_data = book.sheet_by_index(1) # sheet containing Long Term Debt Data (Quarterly)
Lib_data = book.sheet_by_index(2) # sheet containing Liability Data (Quarterly)
Eqt_data = book.sheet_by_index(3) # sheet containing Equity Data (Quarterly)
Vol_data = book.sheet_by_index(4) # sheet containing Volatility Data (Daily)
Rte_data = book.sheet_by_index(5) # sheet containing 10 Yr Treasury Data (Daily)
Spd_data = book.sheet_by_index(6) # sheet containing Spread Data (Daily)
Bta_data = book.sheet_by_index(7) # sheet containing BetaEM Data (Daily)
Fsp_data = book.sheet_by_index(8) # sheet containing Final Stock Price Data (Daily)
Mkt_data = book.sheet_by_index(9) # sheet containing Market Cap Data (Daily)

###########################################
#
# First process everything that's daily
# And put it in the date=>data map
#
###########################################

dateToData = dict()
dateWeekendBool = dict()
dateList = list()


"""
The first thing we have to do is put in 
the risk-free rate r and the market cap
values as a proxy for equity. These need
to happen first because they're the vars
with the dirtiest data, and will define
the scope of the dates to be used from 
the other variables.
"""

## process the market cap data
dates = Mkt_data.col_slice( colx=0, start_rowx=0 )
nDates = len(dates)

for i in range(0, nDates):
        
    ## Put in mkt cap data    
    date_i_float = Mkt_data.cell_value(i,0)
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime("%A %d. %B %Y")
    date_i_formatted = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime('%m-%d-%Y')
    
    dateToData[ date_i_formatted ] = [ float( Mkt_data.cell_value(i,1) ) ]
    
## process the risk-free rate data
dates_rte = Rte_data.col_slice( colx=0, start_rowx=0 )
nDatesRte = len(dates_rte)

for i in range(0, nDatesRte):
    
    ## Put in rfr data    
    date_i_float = Rte_data.cell_value(i,0)
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime("%A %d. %B %Y")
    date_i_formatted = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime('%m-%d-%Y')


    try:
        
        dateToData[ date_i_formatted ].append( float( Rte_data.cell_value(i,1) ) )


    except KeyError:
        
        continue



"""
The current map is as such:
    date => [E,r]

Now we put in the remaining 
daily variables:
    vol, spread, beta
"""

dates_vol = Vol_data.col_slice( colx=0, start_rowx=0 )
nDatesVol = len(dates_vol)


for i in range(0, nDatesVol):
    
    ## Put in rfr data    
    date_i_float = Vol_data.cell_value(i,0)
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime("%A %d. %B %Y")
    date_i_formatted = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime('%m-%d-%Y')

    try:

        dateToData[ date_i_formatted ].append( float( Vol_data.cell_value(i,1) ) )
        dateToData[ date_i_formatted ].append( float( Spd_data.cell_value(i,1) ) )
        dateToData[ date_i_formatted ].append( float( Bta_data.cell_value(i,1) ) )
    
    except KeyError:
        
        continue
    
    

###################################################
# @ this point, all the daily data
# is in the map. Need to get the quarterlies
# map: date => [ vol, spd, bta, fsp, rte ]
###################################################

## Put in the liability data (STD,LTD maybe)
quarterToSomeData = dict()

quarters = STD_data.col_slice( colx=0, start_rowx=0 )
nEntries_q = len( quarters )

for i in range( 0, nEntries_q ):
    
    quarter_i = str( STD_data.cell_value( i, 0 ) )
    quarter_i_formatted = quarter_i[2] + "." + quarter_i[4:]
    
    quarterToSomeData[ quarter_i_formatted ] = [ float( Lib_data.cell_value( i, 1 ) ) ]
    quarterToSomeData[ quarter_i_formatted ].append( float( STD_data.cell_value( i, 1 ) ) )
    quarterToSomeData[ quarter_i_formatted ].append( float( LTD_data.cell_value( i, 1 ) ) )



###################################################
# @ this point, the quarterToSomeData
# map maps quarter=> [lib, std, ltd]
# we need to go through each date in the 
# dateToData map, and see what quarter it's in. 
# reference this quarter, and put each value in the 
# dateToData map
###################################################

for date in dateToData:
    
    ## get values
    for elm in quarterToSomeData[ dateToQuarter(date) ]:
        
        dateToData[ date ].append( elm )
        

"""
Cleaning data: if any values of the 
dataFrame associated to a date are 
0, then remove them from the map
"""

removeKeyList = list()
for k in dateToData.keys():

    if 0 in dateToData[k]:
        
        removeKeyList.append( k )
        
for elm in removeKeyList:
    
    del dateToData[elm]


#############################################
# Now that all the data is centralized,
# use the paramters to do the computations
#############################################

M = 5.0 # assume
outFile = open('fc_may_11.csv','w')

betaBM_List = list()
d_spreadList = list()
LGD_List = list()
PD_List = list()
for k in dateToData.keys():
    
    E = dateToData[k][0]*1000000
    r = dateToData[k][1]/100.0
    sig_E = dateToData[k][2]/100.0
    spd = dateToData[k][3]
    bta = dateToData[k][4]
    B = dateToData[k][5]*1000000
    STD = dateToData[k][6]*1000000
    LTD = dateToData[k][7]*1000000
    
    N_Moody = STD + 0.5*LTD
    A = E + B
    
    initial_guess = [0.2, N_Moody]
    
    # debug
    """
    print("\n====DEBUG====")
    print("sigE: ", sig_E)
    print("spd_phys: ", spd)
    print("beta_EM: ", bta)
    print("rate: ", r)
    print("Liabilities: ", B)
    print("STD: ", STD)
    print("LTD: ", LTD)
    print("Equity: ", E)
    """
    
    try:
                      
        CalibSigA, CalibN = fsolve( fsolve_function, initial_guess, args=(E, sig_E, r, M, A) )

        # d1/d2 aux functions
        d1 = (math.log(A/CalibN) + ((r + 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))
        d2 = (math.log(A/CalibN) + ((r - 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))    
        
        ## Calculate Beta_{B,M}, and append to list to get average later
        beta_BM = (E*norm.cdf(-1.0*d1)*bta)/float(B*norm.cdf(d1))
        
        print(beta_BM)
        
        betaBM_List.append(beta_BM)
        
        ## Calculate DeltaSpread (and Calibrated Spread)
        d_spread, calib_spread = delta_spread( A, CalibSigA, E, sig_E, CalibN, r, M, spd )
        d_spreadList.append( d_spread )
        
        ## Loss Given Default (LGD), Probability of Default (PD)
        LGD = 1.0 - ( ( math.exp(r*M) * A * norm.cdf(-d1) ) / ( CalibN*norm.cdf(-d2)  ) )
        PD = norm.cdf(-d1)
        
        LGD_List.append( LGD )
        PD_List.append( PD )
        
        outStr = str(beta_BM) + "," + str(d_spread) + "," + str(LGD) + "," + str(PD) + "\n"
        outFile.write( outStr )
        

        
    except ValueError:

        continue


plt.plot(betaBM_List, d_spreadList, 'bo')
plt.show()

#print("avg Beta_BM: ", averageOfList( betaBM_List ) )
#print("max Beta_BM: ", max( betaBM_List) )









