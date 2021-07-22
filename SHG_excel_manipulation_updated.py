# -*- coding: utf-8 -*-
"""
Created on Thu Jul  8 10:46:08 2021

@author: natha
"""
from scipy.stats import ttest_ind
import math
import os
import pandas as pd
import numpy as np

#Ask User Heart and Location to fetch dataset and later create excel sheet with series name
heart = 'H1'#input("Heart #: ")
ltn = 'V1'#input("Location #: ")
con = 'ctrl'#input("Condition of Heart (ex. ctrl, blebb): ")

#Set working directory
os.chdir(r"C:\Users\natha\Desktop\CEMB Summer Program 2021\Discher Lab\Ex-vivo chick heart experiments\Experiments\Karan\SHG Analysis\{c}\{h}\{s}".format(c = con, h = heart, s = heart + ltn ))

#file path
file = r"C:\Users\natha\Desktop\CEMB Summer Program 2021\Discher Lab\Ex-vivo chick heart experiments\Experiments\Karan\SHG Analysis\{c}\{h}\{s}\{c}_{s}_data.xlsx".format(c = con, h = heart, s = heart + ltn )


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
        
                 
del channels['channel_9']

    

#Adding characteristic mean for each ROI (new row), +/- 3 z slices of peak
maxs2match = []
channels35 = {'channel_3': pd.DataFrame(None),'channel_4': pd.DataFrame(None),'channel_7': pd.DataFrame(None),'channel_8': pd.DataFrame(None)}
for key, df in channels.items():
    if any(x in key for x in list(channels35.keys())):
        ROI_avgs = []
        for column in df[['T1', 'T2','T3']]:
                min = int(df[[column]].idxmax() - 3)
                max = int(df[[column]].idxmax() + 3)
                #list of index values from min2max
                min2max = list(range(min, max+1, 1))
                
                #append mean of +/- 3 rows of max in column to roi list only if min is in range, if not append a NaN
                if (min >= 0) & (max <= len(df[[column]])-1):
                      #append mean to channels dictionary
                      ROI_avgs.append(float(df[[column]].iloc[min2max].mean(axis=0)))
                      
                      #set channels35 key equal to new df with ROI name and max value index
                      df1 = pd.DataFrame([int(df[[column]].idxmax())], columns = [column])
                      channels35[key] = pd.concat([channels35[key],df1], ignore_index= False, axis=1)
                else:
                      ROI_avgs.append(math.nan)
                      
                      #set channels35 key equal to new df with ROI name
                      df1 = pd.DataFrame(None, columns = [column])
                      channels35[key] = pd.concat([channels35[key],df1], ignore_index= False, axis=1)
                      
                      
        df1 = pd.DataFrame(ROI_avgs, columns = ['T_avgs'])
        channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)

channels5 = {'channel_1': channels35['channel_3'],'channel_2': channels35['channel_4'],'channel_5': channels35['channel_7'],'channel_6': channels35['channel_8']}

for key, df in channels5.items():
    ROI_avgs = []
    for column in channels[key][['T1', 'T2','T3']]:
            min = int(df.iloc[0][column] - 3)
            max = int(df.iloc[0][column] + 3)
            #list of index values from min2max
            min2max = list(range(min, max+1, 1))
            ROI_avgs.append(float(channels[key][column].iloc[min2max].mean(axis=0)))
     
            

    df1 = pd.DataFrame(ROI_avgs, columns = ['T_avgs'])
    channels[key] = pd.concat([channels[key],df1], ignore_index= False, axis=1)


#Add column T_Favg, average of ROI averages
for key, df in channels.items():
    df1 = pd.DataFrame([df['T_avgs'].mean()], columns=['T_Favg'])
    channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)

#Add column T_SEM, Standard error of ROI averages
for key, df in channels.items():
    df1 = pd.DataFrame([df['T_avgs'].std()/math.sqrt(3)], columns=['T_SEM'])
    channels[key] = pd.concat([df,df1], ignore_index= False, axis=1)





#Make Panda for col chart (920,860nm) and 920,860nm dictionary
channels920_860 = {'channel_1':'channel_5', 'channel_2':'channel_6','channel_3':'channel_7','channel_4':'channel_8'}
blank4x5 = [[None,None,None,None],[None,None,None,None],[None,None,None,None],[None,None,None,None],[None,None,None,None]]
colChart920_860data = pd.DataFrame(blank4x5, columns= ['5% Bkwd','5% Fwd','35% Bkwd','35% Fwd'])

colChart920_860data['5% Bkwd'][0] = channels['channel_1']['T_Favg'][0]
colChart920_860data['5% Bkwd'][1] = channels['channel_5']['T_Favg'][0]

colChart920_860data['5% Fwd'][0] = channels['channel_2']['T_Favg'][0]
colChart920_860data['5% Fwd'][1] = channels['channel_6']['T_Favg'][0]

colChart920_860data['35% Bkwd'][0] = channels['channel_3']['T_Favg'][0]
colChart920_860data['35% Bkwd'][1] = channels['channel_7']['T_Favg'][0]

colChart920_860data['35% Fwd'][0] = channels['channel_4']['T_Favg'][0]
colChart920_860data['35% Fwd'][1] = channels['channel_8']['T_Favg'][0]

#SEM
colChart920_860data['5% Bkwd'][2] = channels['channel_1']['T_SEM'][0]
colChart920_860data['5% Bkwd'][3] = channels['channel_5']['T_SEM'][0]

colChart920_860data['5% Fwd'][2] = channels['channel_2']['T_SEM'][0]
colChart920_860data['5% Fwd'][3] = channels['channel_6']['T_SEM'][0]

colChart920_860data['35% Bkwd'][2] = channels['channel_3']['T_SEM'][0]
colChart920_860data['35% Bkwd'][3] = channels['channel_7']['T_SEM'][0]

colChart920_860data['35% Fwd'][2] = channels['channel_4']['T_SEM'][0]
colChart920_860data['35% Fwd'][3] = channels['channel_8']['T_SEM'][0]


#Perform two tailed T-test, unequal variance between 3 averages of T_avgs
colChart920_860data['5% Bkwd'][4] = ttest_ind(channels['channel_1']['T_avgs'][0:3], channels['channel_5']['T_avgs'][0:3], equal_var= False)[1]
colChart920_860data['5% Fwd'][4] = ttest_ind(channels['channel_2']['T_avgs'][0:3], channels['channel_6']['T_avgs'][0:3], equal_var= False)[1]
colChart920_860data['35% Bkwd'][4] = ttest_ind(channels['channel_3']['T_avgs'][0:3], channels['channel_7']['T_avgs'][0:3], equal_var= False)[1]
colChart920_860data['35% Fwd'][4] = ttest_ind(channels['channel_4']['T_avgs'][0:3], channels['channel_8']['T_avgs'][0:3], equal_var= False)[1]
colChart920_860data = colChart920_860data.rename(index= {0:'920nm_T_Favg', 1:'860nm_T_Favg', 2:'920nm_SEM', 3:'860nm_SEM', 4: 'P_value'})

    

#Make Bkwd-Fwd comparison charts
channelsBkwd_Fwd = {'channel_1':'channel_2', 'channel_3':'channel_4','channel_5':'channel_6','channel_7':'channel_8'}
colChartBkwd_Fwd_data = pd.DataFrame(blank4x5, columns= ['5% 920nm','35% 920nm','5% 860nm','35% 860nm'])

colChartBkwd_Fwd_data['5% 920nm'][0] = channels['channel_1']['T_Favg'][0]
colChartBkwd_Fwd_data['5% 920nm'][1] = channels['channel_2']['T_Favg'][0]

colChartBkwd_Fwd_data['35% 920nm'][0] = channels['channel_3']['T_Favg'][0]
colChartBkwd_Fwd_data['35% 920nm'][1] = channels['channel_4']['T_Favg'][0]

colChartBkwd_Fwd_data['5% 860nm'][0] = channels['channel_5']['T_Favg'][0]
colChartBkwd_Fwd_data['5% 860nm'][1] = channels['channel_6']['T_Favg'][0]

colChartBkwd_Fwd_data['35% 860nm'][0] = channels['channel_7']['T_Favg'][0]
colChartBkwd_Fwd_data['35% 860nm'][1] = channels['channel_8']['T_Favg'][0]

#SEM
colChartBkwd_Fwd_data['5% 920nm'][2] = channels['channel_1']['T_SEM'][0]
colChartBkwd_Fwd_data['5% 920nm'][3] = channels['channel_2']['T_SEM'][0]

colChartBkwd_Fwd_data['35% 920nm'][2] = channels['channel_3']['T_SEM'][0]
colChartBkwd_Fwd_data['35% 920nm'][3] = channels['channel_4']['T_SEM'][0]

colChartBkwd_Fwd_data['5% 860nm'][2] = channels['channel_5']['T_SEM'][0]
colChartBkwd_Fwd_data['5% 860nm'][3] = channels['channel_6']['T_SEM'][0]

colChartBkwd_Fwd_data['35% 860nm'][2] = channels['channel_7']['T_SEM'][0]
colChartBkwd_Fwd_data['35% 860nm'][3] = channels['channel_8']['T_SEM'][0]


#Perform two tailed T-test, unequal variance between 3 averages of T_avgs
colChartBkwd_Fwd_data['5% 920nm'][4] = ttest_ind(channels['channel_1']['T_avgs'][0:3], channels['channel_2']['T_avgs'][0:3], equal_var= False)[1]
colChartBkwd_Fwd_data['35% 920nm'][4] = ttest_ind(channels['channel_3']['T_avgs'][0:3], channels['channel_4']['T_avgs'][0:3], equal_var= False)[1]
colChartBkwd_Fwd_data['5% 860nm'][4] = ttest_ind(channels['channel_5']['T_avgs'][0:3], channels['channel_6']['T_avgs'][0:3], equal_var= False)[1]
colChartBkwd_Fwd_data['35% 860nm'][4] = ttest_ind(channels['channel_7']['T_avgs'][0:3], channels['channel_8']['T_avgs'][0:3], equal_var= False)[1]
colChartBkwd_Fwd_data = colChartBkwd_Fwd_data.rename(index= {0:'Bkwd', 1:'Fwd', 2:'Bkwd_SEM', 3:'Fwd_SEM', 4: 'P_value'})


#Ratio panda (920,860nm)
colChart920v860_Rdata = {'5% Bkwd': pd.DataFrame(None),'5% Fwd': pd.DataFrame(None),'35% Bkwd': pd.DataFrame(None),'35% Fwd': pd.DataFrame(None)}
for column in channels['channel_1'][['T1','T2','T3']]:
    colChart920v860_Rdata['5% Bkwd'][column]= channels['channel_1'][column]/channels['channel_5'][column]
colChart920v860_Rdata['5% Bkwd'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])

for column in channels['channel_2'][['T1','T2','T3']]:
    colChart920v860_Rdata['5% Fwd'][column]= channels['channel_2'][column]/channels['channel_6'][column]
colChart920v860_Rdata['5% Fwd'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])
    
for column in channels['channel_3'][['T1','T2','T3']]:
    colChart920v860_Rdata['35% Bkwd'][column]= channels['channel_3'][column]/channels['channel_7'][column]
colChart920v860_Rdata['35% Bkwd'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])
    
for column in channels['channel_4'][['T1','T2','T3']]:
    colChart920v860_Rdata['35% Fwd'][column]= channels['channel_4'][column]/channels['channel_8'][column]
colChart920v860_Rdata['35% Fwd'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])

for key, df in colChart920v860_Rdata.items():
    #remove -inf and infs (less than -1B and greater than 1B )and replace with NAN
    colChart920v860_Rdata[key] = df.mask(df < -1000000000.0 )
    colChart920v860_Rdata[key] = colChart920v860_Rdata[key].mask(df > 1000000000.0 )


#Ratio panda (Fwd, Bkwd)
colChartFwdvBkwd_Rdata = {'5% 920nm': pd.DataFrame(None),'35% 920nm': pd.DataFrame(None),'5% 860nm': pd.DataFrame(None),'35% 860nm': pd.DataFrame(None)}
for column in channels['channel_1'][['T1','T2','T3']]:
    colChartFwdvBkwd_Rdata['5% 920nm'][column]= channels['channel_1'][column]/channels['channel_2'][column]
colChartFwdvBkwd_Rdata['5% 920nm'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])

for column in channels['channel_2'][['T1','T2','T3']]:
    colChartFwdvBkwd_Rdata['35% 920nm'][column]= channels['channel_3'][column]/channels['channel_4'][column]
colChartFwdvBkwd_Rdata['35% 920nm'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])
    
for column in channels['channel_3'][['T1','T2','T3']]:
    colChartFwdvBkwd_Rdata['5% 860nm'][column]= channels['channel_5'][column]/channels['channel_6'][column]
colChartFwdvBkwd_Rdata['5% 860nm'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])
    
for column in channels['channel_4'][['T1','T2','T3']]:
    colChartFwdvBkwd_Rdata['35% 860nm'][column]= channels['channel_7'][column]/channels['channel_8'][column]
colChartFwdvBkwd_Rdata['35% 860nm'].insert(column = 'Depth', loc = 0, value = channels['channel_1']['Depth'])

for key, df in colChartFwdvBkwd_Rdata.items():
    #remove -inf and infs (less than -1B and greater than 1B )and replace with NAN
    colChartFwdvBkwd_Rdata[key] = df.mask(df < -1000000000.0 )
    colChartFwdvBkwd_Rdata[key] = colChartFwdvBkwd_Rdata[key].mask(df > 1000000000.0 )
        

##########
##########
##########
##########

#make Excel writer with user input for file name and workbook object
writer = pd.ExcelWriter('E5_{s}_{c}_FA_fixed_analysis.xlsx'.format(s = heart + ltn, c = con))
workbook = writer.book

##########################################################################
#Create sheet for Fwd,Bkwd col chart, can't do Chartsheet bc of p-values
colChartBkwd_Fwd_data.to_excel(excel_writer = writer, sheet_name = 'FwdvsBkwd' , index = True)
chartsheet0 =writer.sheets['FwdvsBkwd']

#create chart object for 920,860nm charts
colChartBkwd_Fwd = workbook.add_chart({'type': 'column'})

#save length of values length in colChartBkwd_Fwd_data panda
col_y_Bkwd_Fwd = len(colChartBkwd_Fwd_data.columns)

#Add 920 series from colChartBkwd_Fwd_data to colChartBkwd_Fwd object
colChartBkwd_Fwd.add_series({
    'name':       'Bkwd',
    'categories': ['FwdvsBkwd', 0, 1, 0, col_y_Bkwd_Fwd],
    'values':     ['FwdvsBkwd', 1, 1, 1, col_y_Bkwd_Fwd],
    'y_error_bars': {
        'type':         'custom',
        'plus_values':  list(colChartBkwd_Fwd_data.iloc[2]),
        'minus_values': list(colChartBkwd_Fwd_data.iloc[2]),}
})

#Add 860 series from colChartBkwd_Fwd_data to colChartBkwd_Fwd object
colChartBkwd_Fwd.add_series({
    'name':       'Fwd',
    'categories': ['FwdvsBkwd', 0, 1, 0, col_y_Bkwd_Fwd],
    'values':     ['FwdvsBkwd', 2, 1, 2, col_y_Bkwd_Fwd],
    'y_error_bars': {
        'type':         'custom',
        'plus_values':  list(colChartBkwd_Fwd_data.iloc[3]),
        'minus_values': list(colChartBkwd_Fwd_data.iloc[3]),}
})  

# Set name on axis of colChartBkwd_Fwd object and insert to 920vs860nm sheet
colChartBkwd_Fwd.set_title({'name': 'FwdvsBkwd'})
colChartBkwd_Fwd.set_x_axis({'name': 'Groups'})
colChartBkwd_Fwd.set_y_axis({'name': 'SHG signal',
                  'major_gridlines': {'visible': True}})

#Insert chartsheet
chartsheet0.insert_chart('G2', colChartBkwd_Fwd)


#######################################################################
#Create sheet for 920,860nm col chart, can't do Chartsheet bc of p-values
colChart920_860data.to_excel(excel_writer = writer, sheet_name = '920vs860nm' , index = True)
chartsheet1 =writer.sheets['920vs860nm']

#create chart object for 920,860nm charts
colChart920_860 = workbook.add_chart({'type': 'column'})

#save length of values length in colChart920_860data panda
col_y_920v860 = len(colChart920_860data.columns)

#Add 920 series from colChart920_860data to colChart920_860 object
colChart920_860.add_series({
    'name':       '920nm',
    'categories': ['920vs860nm', 0, 1, 0, col_y_920v860],
    'values':     ['920vs860nm', 1, 1, 1, col_y_920v860],
    'y_error_bars': {
        'type':         'custom',
        'plus_values':  list(colChart920_860data.iloc[2]),
        'minus_values': list(colChart920_860data.iloc[2]),}
})

#Add 860 series from colChart920_860data to colChart920_860 object
colChart920_860.add_series({
    'name':       '860nm',
    'categories': ['920vs860nm', 0, 1, 0, col_y_920v860],
    'values':     ['920vs860nm', 2, 1, 2, col_y_920v860],
    'y_error_bars': {
        'type':         'custom',
        'plus_values':  list(colChart920_860data.iloc[3]),
        'minus_values': list(colChart920_860data.iloc[3]),}
})  

# Set name on axis of colChart920_860 object and insert to 920vs860nm sheet
colChart920_860.set_title({'name': '920vs860nm'})
colChart920_860.set_x_axis({'name': 'Groups'})
colChart920_860.set_y_axis({'name': 'SHG signal',
                  'major_gridlines': {'visible': True}})

#Insert chartsheet
chartsheet1.insert_chart('G2', colChart920_860)

#################################################
#Create chartsheet for 920v860 data ratios

#Insert Ratio 960v860nm into excel
for key, df in colChart920v860_Rdata.items():
    df.to_excel(excel_writer = writer, sheet_name = str(key), index = True)
    
    # Make worksheet object for each channel in channels dictionary
    worksheet = writer.sheets[str(key)]
    
    # Create a scatter chart object, FOV-channel specific chart, int. vs depth
    scatterChart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    for column in df[['T1', 'T2','T3']]:
        # Get the number of rows and column index
        max_row = len(df)
        col_x = df.columns.get_loc('Depth') + 1
        col_y = df.columns.get_loc(column) + 1
        
        # Create the scatter plot with error cols if T_SEM is equal to something, if not don't add error
        scatterChart.add_series({
            'name':       column,
            'categories': [str(key), 1, col_x, max_row, col_x],
            'values':     [str(key), 1, col_y, max_row, col_y],
            'marker':     {'type': 'circle', 'size': 4},
        })
            
    # Set name on axis
    scatterChart.set_title({'name': key})
    scatterChart.set_x_axis({'name': 'Depth'})
    scatterChart.set_y_axis({'name': 'SHG signal',
                      'major_gridlines': {'visible': False}})
    
    # Insert the charts into the worksheet in field D2
    worksheet.insert_chart('I10', scatterChart)

#####################################
#create sheets from channels dictionary and populate with channel 
for key, df in channels.items():
    df.to_excel(excel_writer = writer, sheet_name = str(key), index = True)
    
    # Make worksheet object for each channel in channels dictionary
    worksheet = writer.sheets[str(key)]
    
    # Create a scatter chart object, FOV-channel specific chart, int. vs depth
    scatterChart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    for column in df[['T1', 'T2','T3']]:
        # Get the number of rows and column index
        max_row = len(df)
        col_x = df.columns.get_loc('Depth') + 1
        col_y = df.columns.get_loc(column) + 1
        
        # Create the scatter plot with error cols if T_SEM is equal to something, if not don't add error
        if math.isnan(df['T_SEM'].iloc[0]):       
            scatterChart.add_series({
                'name':       column,
                'categories': [str(key), 1, col_x, max_row, col_x],
                'values':     [str(key), 1, col_y, max_row, col_y],
                'marker':     {'type': 'circle', 'size': 4},
            })            
        else:
            scatterChart.add_series({
                'name':       column,
                'categories': [str(key), 1, col_x, max_row, col_x],
                'values':     [str(key), 1, col_y, max_row, col_y],
                'marker':     {'type': 'circle', 'size': 4},
                'y_error_bars': {'type': 'standard_error'}#'value': df['T_SEM'].iloc[0],},
            })
            
    # Set name on axis
    scatterChart.set_title({'name': key})
    scatterChart.set_x_axis({'name': 'Depth'})
    scatterChart.set_y_axis({'name': 'SHG signal',
                      'major_gridlines': {'visible': False}})
    
    # Insert the charts into the worksheet in field D2
    worksheet.insert_chart('I10', scatterChart)
        
#close excel file       
writer.close()
