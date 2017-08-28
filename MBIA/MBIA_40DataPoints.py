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

def ListAverage( L ):
    
    return sum(L)/float(len(L))
    
    
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
    d2 = (math.log(A/N) + (r - 0.5*(sigma_A**2)))/(sigma_A*math.sqrt(M))
    
    #return
    #out = A*norm.cdf(d1) - N*math.exp(-1*(r*M))*norm.cdf(d2) - E_market
    #out = sigma_E_market - sigma_A*A*norm.cdf(d1)
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
book = xlrd.open_workbook( "MBIA_1factor.xlsx" )
STD_data = book.sheet_by_index(0) # sheet containing Short Term Debt Data (Quarterly)
LTD_data = book.sheet_by_index(1) # sheet containing Long Term Debt Data (Quarterly)
Lib_data = book.sheet_by_index(2) # sheet containing Liability Data (Quarterly)
Eqt_data = book.sheet_by_index(3) # sheet containing Equity Data (Quarterly)
Vol_data = book.sheet_by_index(4) # sheet containing Volatility Data (Daily)
Rte_data = book.sheet_by_index(5) # sheet containing 10 Yr Treasury Data (Daily)
Spd_data = book.sheet_by_index(6) # sheet containing Spread Data (Daily)
Bta_data = book.sheet_by_index(7) # sheet containing BetaEM Data (Daily)
Fsp_data = book.sheet_by_index(8) # sheet containing Final Stock Price Data (Daily)

"""
###########################################
#
# First process everything that's daily
# And put it in the date=>data map
#
###########################################

dateToData = dict()
dateWeekendBool = dict()
dateList = list()

## process volatility data
dates = Vol_data.col_slice( colx=0, start_rowx=0 )
nEntries = len(dates)

for i in range( 0, nEntries ):
    
    ## Put in volatility data    
    date_i_float = Vol_data.cell_value(i,0)
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime("%A %d. %B %Y")
    date_i_formatted = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime('%d-%m-%Y')
    day = date_i.split(" ")[0]
    
    dateList.append( date_i_formatted )
    
    if notWeekendCheck( day ):
        
        dateWeekendBool[ date_i_formatted ] = True

        dateToData[ date_i_formatted ] = [ day ]
        dateToData[ date_i_formatted ].append( float( Vol_data.cell_value( i,1 ) ) )
        dateToData[ date_i_formatted ].append( float( Spd_data.cell_value( i,1 ) ) )
        dateToData[ date_i_formatted ].append( float( Bta_data.cell_value( i,1 ) ) )
        dateToData[ date_i_formatted ].append( float( Fsp_data.cell_value( i,1 ) ) )
        
    else:
        
        dateWeekendBool[ date_i_formatted ]= False
    
"""



###########################################
#
# First process everything that's quarterly
# And put it in the quarter=>data map, q2d.
# Vars: E, B, STD, LTD
#
###########################################

q2d = dict()
quarters = Eqt_data.col_slice( colx=0, start_rowx=0 )

for i in range (0, len(quarters)):
    
    q = Eqt_data.cell_value( i, 0 )
    temp = q.strip().split(" ")
    quarter = str(temp[0][2]) + "." + str(temp[1])
    
    q2d[quarter] = [ Eqt_data.cell_value( i, 1 ) ]
    q2d[quarter].append( Lib_data.cell_value( i, 1 ) )
    q2d[quarter].append( STD_data.cell_value( i ,1 ) )
    q2d[quarter].append( LTD_data.cell_value( i ,1 ) )
    
    
###########################################
#
# Currently:
#   q2d[ quarter ] = [E, B, STD, LTD]
#
# Now, we must average and put in the
# daily variables:   
#   sigma, Beta_{E,M}, S^{Phy}, r
#
###########################################

## set up a date to quarter map
date2quarterMap = dict()
dates = Vol_data.col_slice( colx=0, start_rowx=0 )

newMap = dict()
for i in range( 0, len(dates) ):
    
    date_i_float = Vol_data.cell_value(i,0)
    date_i = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime("%A %d. %B %Y")
    date_i_formatted = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float, book.datemode ) ).strftime('%d-%m-%Y')
    day = date_i.split(" ")[0]
    
    if notWeekendCheck( day ):
    
        try:
            
            newMap[ dateToQuarter( date_i_formatted ) ][0].append( Vol_data.cell_value( i,1 ) )
            newMap[ dateToQuarter( date_i_formatted ) ][1].append( Bta_data.cell_value( i,1 ) )
            newMap[ dateToQuarter( date_i_formatted ) ][2].append( Spd_data.cell_value( i,1 ) )
        
        except KeyError:
            
            newMap[ dateToQuarter( date_i_formatted ) ] = [[ float(Vol_data.cell_value(i,1)) ], \
                   [ float(Bta_data.cell_value(i,1)) ],[ float(Spd_data.cell_value(i,1)) ] ] 
        







