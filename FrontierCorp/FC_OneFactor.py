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
    d2 = (math.log(A/N) + (r - 0.5*(sigma_A**2)))/(sigma_A*math.sqrt(M))
    
    #return
    #out = A*norm.cdf(d1) - N*math.exp(-1*(r*M))*norm.cdf(d2) - E_market
    #out = sigma_E_market - sigma_A*A*norm.cdf(d1)
    out = [A*norm.cdf(d1) - N*math.exp(-1*(r*M))*norm.cdf(d2) - E_market]
    out.append(sigma_A*A*norm.cdf(d1) - sigma_E_market*E_market)

    return out


def delta_spread( A, calib_sigma_A, E_market, sigma_E_market, calibN, r, M, spd ):
    
    d1 = (math.log(A/calibN) + ((r + 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))
    d2 = (math.log(A/calibN) + ((r - 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))    
    
    spread_calib = (-1.0/M)*math.log( norm.cdf(d2) + ((A*math.exp(r*M))/calibN)*norm.cdf(-1.0*d1) )
    spread_calib *= 10000.0
    
    #return (spd-spread_calib2)
    return (spd-spread_calib)


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
    

## Put in the treasury data
dates_Rte = Rte_data.col_slice( colx=0, start_rowx=0 )
nEntries_Rte = len( dates_Rte )
for i in range( 0, nEntries_Rte ):
    
    date_i_float_rte = Rte_data.cell_value(i,0)
    date_i_formatted_rte = datetime.datetime( *xlrd.xldate_as_tuple( date_i_float_rte, book.datemode ) ).strftime('%d-%m-%Y')    
    dateToData[ date_i_formatted_rte ].append( float( Rte_data.cell_value( i,1 ) ) )


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
        

###################################################
# @ this point, the dateToData map contains
# everything except Equity, the special case.
# Set up the equity map first, 
# quarter to Equity values
###################################################

## Put in the Equity data
quarterToEquity = dict()

quarters_Eqt = Eqt_data.col_slice( colx=0, start_rowx=0 )
nEntries_Eqt = len( quarters_Eqt )

for i in range( 0, nEntries_Eqt ):
        
    quarter_i = str( Eqt_data.cell_value( i, 0 ) )
    quarter_i_formatted = quarter_i[2] + "." + quarter_i[4:]
    
    quarterToEquity[ quarter_i_formatted ] = float( Eqt_data.cell_value( i, 1 ) )


###################################################
# @ this point, the quarterToEquity map
# is set up. Set up quarter=>#ofShares map
###################################################

## Put in number shares
quarterToNumShares = dict()
                         
for i in range( 0, len(dateList) ):
    
    date = dateList[i]
    
    weekend_check = dateWeekendBool[date]
    day_month = date[:-5]
    
    if firstDayCheck( day_month ) and (weekend_check==True):
        
        quarter = dateToQuarter( date )
        fsp = dateToData[ date ][4]
        Equity_val = quarterToEquity[ quarter ]
        numShares = Equity_val/float(fsp)

        # insert into quarterToNumShares map
        quarterToNumShares[ quarter ] = numShares
                          
    if firstDayCheck( day_month ) and (weekend_check==False):
        
        temp_i = i #save the location of i
        while( dateWeekendBool[ dateList[i] ] == False ):
            
            i+=1
            
        date_i = dateList[i] #save the next non-weekend date
        i = temp_i #refresh to true position of i
        
        quarter = dateToQuarter( date_i )
        fsp = dateToData[ date_i ][4]
        Equity_val = quarterToEquity[ quarter ]
        numShares = Equity_val*1000000/float(fsp)

        # insert into quarterToNumShares map
        quarterToNumShares[ quarter ] = numShares
            

###################################################
# @ this point, the quarterToNumShares map
# is set up. Now we just need to put Equity
# into the dateToData map
###################################################

for date in dateToData:
    
    quarter = dateToQuarter( date )
    numShares = quarterToNumShares[ quarter ]
    
    dateToData[date].append(numShares)
    

#print(dateToData)

###################################################
# @ this point, everything is in the dataFrame
# just need to do math
###################################################
    
#############################################
# Now that all the data is centralized,
# use the paramters to do the computations
#############################################


A_init = 70000000
sigmaA_init = 0.2
init_guess = [A_init, sigmaA_init]
M = 5.0 # assume


outFile = open('FC_testingFile.csv','w')
betaBM_List = list()
for k in dateToData.keys():
        
    #define input parameters for calibration
    sig_E = dateToData[k][1]/100.0
    spd = dateToData[k][2]
    bta = dateToData[k][3]
    fsp = dateToData[k][4]
    r = dateToData[k][5]/100.0
    B = dateToData[k][6]*1000000
    STD = dateToData[k][7]*1000000
    LTD = dateToData[k][8]*1000000
    numShares = dateToData[k][9]
    E = float(fsp*numShares)*1000000
    N_Moody = STD+0.5*LTD    
    N_Liab = B
    A = B + E
    
    initial_guess = [0.2, N_Moody]

    if (spd != 0.0):
        
        try:
              
            CalibSigA, CalibN = fsolve( fsolve_function, initial_guess, args=(E, sig_E, r, M, A) )
            
            """            
            print("\n====DEBUG====")
            print("sigE: ", sig_E)
            print("spd_phys: ", spd)
            print("beta_EM: ", bta)
            print("fStockPrice: ", fsp)
            print("rate: ", r)
            print("Liabilities: ", B)
            print("STD: ", STD)
            print("LTD: ", LTD)
            print("N_Moodys: ", N_Moody)
            print("N_Moodys/B: ", N_Moody/float(B))
            print("Equity: ", E)
            print("CalibSigA: ", CalibSigA)
            print("CalibN: ", CalibN)
            print("B/CalibN: ", B/float(CalibN))
            """
            
            ## Get d1/d2
            d1 = (math.log(A/CalibN) + ((r + 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))
            d2 = (math.log(A/CalibN) + ((r - 0.5*(CalibSigA**2))*M))/(CalibSigA*math.sqrt(M))    
            
            beta_BM = (E*norm.cdf(-1.0*d1)*bta)/float(B*norm.cdf(d1))
            phi_d1 = norm.cdf(d1)
            
#            print("\n====DEBUG====")
#            print("CalibSigA: ", CalibSigA)
#            print("CalibN: ", CalibN)
#            print("beta_BM: ", beta_BM)
            
            betaBM_List.append(beta_BM)
        
            ## Calculate S^Physical - S^Calibrate
            d_spread = delta_spread( A, CalibSigA, E, sig_E, CalibN, r, M, spd )
            
            #print("deltaSpread: ", d_spread)

            outStr = str(beta_BM) + "," + str(d_spread) + "\n"
            outFile.write( outStr )

            """
            ## Get LGD / Beta_BM (Beta_EM*constants)
            LGD = 1.0 - ( ( math.exp(r*M) * A * norm.cdf(-d1) ) / ( N*norm.cdf(-d2)  ) )
            PD = norm.cdf(-d1)
            #bta_BM = ( (norm.cdf(-d1)*E*1000000*bta) / (float(B*norm.cdf(d1))) )
            bta_BM = ( (norm.cdf(-d1)*E*1000000*bta) / (float(B*norm.cdf(d1))) )
            #print("Beta_BM: ", bta_BM)
            
#            print("\n SigA: ", CalibSigA)
#            print("Beta_BM: ", bta_BM)

            
            beta_BM_list.append(bta_BM)
            beta_EM_list.append(bta)
            
            ## Output, all files
            out_str = str(bta_BM) + "," + str(bta) + "," + str(d_spread) + "\n"
            fc_dspreadFile2.write( out_str )
            
            out_str2 = str(bta) + "," + str(LGD) + "," + str(PD) + "\n"
            fc_PdFile.write( out_str2 )
            
            out_str3 = k + "," + str(PD) + "," + str(spd) + "\n"
            fc_PdSpread.write( out_str3 )
            
            out_str4 = str(PD) + "," + str(LGD) + "\n"
            last_doc.write( out_str4 )
            
            out_str5 = str(cnt) + "," + str(PD) + "\n"
            fc_TimeSeries.write( out_str5 )
            
            out_str6 = str(E) + "," + str(spd) + "\n"
            fc_betaEBFile.write( out_str6 )
            
            out_str7 = str(CalibSigA) + "," + str(bta_BM) + "\n"
            sigA_BtaBM.write( out_str7 )
            
            #count_list.append(cnt)
            cnt+=1
            
            """
            
        except ValueError:
            
            continue
            
    

#print("\nbetaBM_List_avg: ", sum(betaBM_List)/float(len(betaBM_List)))

#fc_dspreadFile2.close()
#fc_PdFile.close()
#fc_PdSpread.close()
#last_doc.close()
#fc_betaEBFile.close()
