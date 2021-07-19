# -*- coding: utf-8 -*-
"""
Created on Thu Jul  8 10:46:08 2021

@author: natha
"""

import math
import os
import pandas as pd

#Ask User Heart and Location to fetch dataset and later create excel sheet with series name
heart = 'H1'#input("Heart #: ")
ltn = 'V1'#input("Location #: ")
con = 'ctrl'#input("Condition of Heart (ex. ctrl, blebb): ")

#Set working directory
os.chdir(r"C:\Users\natha\Desktop\CEMB Summer Program 2021\Discher Lab\Ex-vivo heart experiments\Karan\SHG Analysis\{c}\{h}\{s}".format(c = con, h = heart, s = heart + ltn ))

file = r"C:\Users\natha\Desktop\CEMB Summer Program 2021\Discher Lab\Ex-vivo heart experiments\Karan\SHG Analysis\{c}\{h}\{s}\{c}_{s}_data.xlsx".format(c = con, h = heart, s = heart + ltn )
#Data must have labels T1, T2, T3, B1, B2
#Must be read in from imagej plugin, time series analyzer
#Read Time Series Analyzer Data into Panda
dataO = pd.read_excel(file, usecols = "A:E")


#Make single column panda of average of two backgrounds
dataBavg = pd.DataFrame(dataO[['B1', 'B2']].mean(axis=1), columns = ['B_avg'])

#Make panda of each tissue ROI signal corrected(tissue mean - background average)
dataT = dataO[ ["T1", "T2", "T3"] ].sub(dataBavg['B_avg'], axis=0)


#Add channel column, recreates channelwise parsing of image to measure data points
numCh = 9
Z_slices = 24#int(input('Number of Z slices: '))
total = numCh*Z_slices
rows = []

j=1
for i in range(total):
    rows.append(j)
    j = j+1
    if j == numCh+1:
        j=1

dataT['Ch'] = rows
    


#Create channel dictionary with dynamic keys, final channel dictionary has index sorted data
channels = {}
for i in range(1, numCh+1):
    key = 'channel_'+str(i)    

    df = dataT.loc[dataT['Ch'] == i].sort_index(axis = 0)
    channels[key] = df


#make a list of depth (increasing by 5) for each channel in channels dictionary 
#and insert at beginning of df and reset indexes
for df in channels.values():
    depth = list(range(0,len(df)*5,5))
    df.insert(0,'Depth',depth)
    df.reset_index(inplace = True, drop = True)
        
                 


    

#Adding characteristic mean for each ROI (new row), +/- 3 z slices of peak
maxs2match = []
channels35 = {'3': None,'4': None,'7': None,'8': None}
for key, df in channels.items():
    ROI_avgs = []
    for column in df[['T1', 'T2','T3']]:
            min = int(df[[column]].idxmax() - 3)
            max = int(df[[column]].idxmax() + 3)
            #list of index values from min2max
            min2max = list(range(min, max+1, 1))
            
            #append mean of +/- 3 rows of max in column to roi list only if min is in range, if not append a NaN
            if (min >= 0) & (max <= len(df[[column]])-1):
                  ROI_avgs.append(float(df[[column]].iloc[min2max].mean(axis=0)))
                
            else:
                  ROI_avgs.append(math.nan)
                  
    df1 = pd.DataFrame(ROI_avgs, columns = ['T_avgs'])
    channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)


#Add column T_Favg, average of ROI averages
for key, df in channels.items():
    df1 = pd.DataFrame([df['T_avgs'].mean()], columns=['T_Favg'])
    channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)

#Add column T_SEM, Standard error of ROI averages
for key, df in channels.items():
    df1 = pd.DataFrame([df['T_avgs'].std()/math.sqrt(3)], columns=['T_SEM'])
    channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)


##########
##########
##########
##########
'''
#make Excel writer with user input for file name
writer = pd.ExcelWriter('E5_{s}_{c}_FA_fixed_analysis.xlsx'.format(s = heart + ltn, c = con))

#create sheets from channels dictionary and populate with channel 
for key, df in channels.items():
    df.to_excel(excel_writer = writer, sheet_name = str(key), index = True)
    
    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = writer.book
    worksheet = writer.sheets[str(key)]
    
    
    # Create a scatter chart object.
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    for column in df[['T1', 'T2','T3']]:
        # Get the number of rows and column index
        max_row = len(df)
        col_x = df.columns.get_loc('Depth') + 1
        col_y = df.columns.get_loc(column) + 1
        
        # Create the scatter plot with error bars if T_SEM is equal to something, if not don't add error
        if math.isnan(df['T_SEM'].iloc[0]):       
            chart.add_series({
                'name':       column,
                'categories': [str(key), 1, col_x, max_row, col_x],
                'values':     [str(key), 1, col_y, max_row, col_y],
                'marker':     {'type': 'circle', 'size': 4},
            })            
        else:
            chart.add_series({
                'name':       column,
                'categories': [str(key), 1, col_x, max_row, col_x],
                'values':     [str(key), 1, col_y, max_row, col_y],
                'marker':     {'type': 'circle', 'size': 4},
                'y_error_bars': {'type': 'percentage','value': df['T_SEM'].iloc[0],},
            })
            
    # Set name on axis
    chart.set_title({'name': key})
    chart.set_x_axis({'name': 'Depth'})
    chart.set_y_axis({'name': 'SHG signal',
                      'major_gridlines': {'visible': False}})
    
    # Insert the chart into the worksheet in field D2
    worksheet.insert_chart('I10', chart)
    
    
    
    
    
    
writer.save()
writer.close()
'''