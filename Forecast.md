# Forecast
#Forecast Model
import pandas as pd
import numpy as np
from datetime import datetime

#Load Import File
data_fp = str(input("Enter data path: /Users/bkanchipurambaburavi/Downloads - Harness Forecast Tester.xlsx"))
data_file = pd.ExcelFile(data_fp)
tsla_forecast = pd.read_excel(data_file, sheet_name='Tesla Forecast')
tsla_forecast = tsla_forecast.drop(axis=1,columns="Update Date")
#Changing the datetime format for time series data to short date format
for i in range(11,115):
    tsla_forecast = tsla_forecast.rename(columns={tsla_forecast.columns[i]: tsla_forecast.columns[i].strftime("%m/%d/%Y")})
#Reading b_h and comp sheet
b_h = pd.read_excel(data_file, sheet_name='Harnesses')
comp = pd.read_excel(data_file, sheet_name='Components')
print("Ok")

'''/Users/bkanchipurambaburavi/Downloads/HFTNEWABI.xlsx'''

#Clean b_h sheet
b_h = b_h.drop(columns='Coax?')
b_h_col_rename = {'Harness/TLA PN':'Harness/PCBA PN'}
b_h = b_h.rename(columns=b_h_col_rename)

#Clean Components Sheet, nothing to put here yet
comp_col_rename = {'Tesla PN':'TPN', 'Harness PN':'Harness/PCBA PN', 'BOM UOM':'UOM'}
comp = comp.rename(columns=comp_col_rename)

#manipulating the split date format of date intervals to a specific entity for comparison with status change date


for i in range(len(b_h['Status Change Date'])):
    if (b_h.iloc[i,14] == str('8/1/2022 - 10/1/2022')) is True:
        b_h.iloc[i,14] = str('2015-01-01 00:00:00')
    else:
        b_h.iloc[i,14] = b_h.iloc[i,14]



#Changing the datetime format for time series data to short date format
b_h['Status Change Date'] = pd.to_datetime(b_h['Status Change Date'])
b_h['Status Change Date'] = b_h['Status Change Date'].dt.strftime('%m/%d/%Y')

print(b_h['Status Change Date'])

#Merge Tesla Forecast with b_h
b_h = pd.merge(b_h,tsla_forecast,how='left')
print(b_h.shape)
print(b_h.head())

#Parse forecast qty columns
print(list(b_h.columns))
num_col_start = str(input("Please enter last column name before forecast data for boards and harness sheet: "))
num_col_start_indx = b_h.columns.get_loc(num_col_start) + 1
b_h_qtys = b_h.iloc[:, num_col_start_indx:]
print(b_h_qtys.head())

#Multiply forecast qty by QPV
b_h['QPV'] = pd.to_numeric(b_h['QPV'], downcast='float')
b_h_forecast = b_h_qtys.mul(b_h['QPV'],axis=0)
print(b_h_forecast.head())


#Update board and harness forecast numbers to reflect multiplied QPV
b_h.iloc[:, num_col_start_indx:] = b_h_forecast
print(b_h.head())

#Offset Updation
b_h.iloc[:,27:] = b_h.apply(lambda x: x.iloc[27:].shift(-x.loc['Offset']),axis=1,result_type='expand')
b_h = b_h.fillna(value=None, method='ffill', axis=1, inplace=False, limit=None, downcast=None)
print(b_h.head())
#print(b_h.shape)
print("Done")

#Based on comparing the status change date and time series date, multipying the respective by current or future QPV

for column in range(26,130):
    for row in range(len(b_h['Status Change Date'])):
        if (datetime.strptime(b_h['Status Change Date'][row], '%m/%d/%Y') == pd.to_datetime("01/01/2015")) is True:
            if datetime.strptime(b_h.columns[column], '%m/%d/%Y') < datetime.strptime(str("08/01/2022"), '%m/%d/%Y'):
                b_h.iloc[row,column] = 0
            if datetime.strptime(b_h.columns[column], '%m/%d/%Y') > datetime.strptime(str("10/01/2022"), '%m/%d/%Y'): 
                b_h.iloc[row,column] = 0
            else: 
                b_h.iloc[row,column] = b_h.iloc[row,column]*b_h.iloc[row,10]
        else:
            if (datetime.strptime(b_h['Status Change Date'][row], '%m/%d/%Y') > datetime.strptime(b_h.columns[column], '%m/%d/%Y')) is True:
                b_h.iloc[row,column] = b_h.iloc[row,column]*b_h.iloc[row,10]
            else:
                b_h.iloc[row,column] = b_h.iloc[row,column]*b_h.iloc[row,11]
            
            
print(b_h.iloc[:,14])

for i in range(len(b_h['Status Change Date'])):
    if (datetime.strptime(b_h['Status Change Date'][i], '%m/%d/%Y') == pd.to_datetime("01/01/2015")) is True:
        b_h.iloc[i,14] = str('8/1/2022 - 10/1/2022')        
    else:
        b_h.iloc[i,14] = b_h.iloc[i,14]
        
        
#pivot summary after offset updation
for i in range (0,4): 
    b_h.iloc[:, 16+i] = b_h.iloc[:,26 + 13*i : 26 + 13*(i+1)].sum(axis = 1)

for i in range (0,4): 
    b_h.iloc[:, 21+i] = b_h.iloc[:,78 + 13*i : 78 + 13*(i+1)].sum(axis = 1)

#quarterly and yearly pivot summary    
b_h.iloc[:, 20] = b_h.iloc[:, 16:20].sum(axis = 1) 
b_h.iloc[:, 25] = b_h.iloc[:, 21:25].sum(axis = 1)

print(b_h)
print("Done")

#Change b_h sheet column names for comp sheet
b_h_rename = {'Description':'Harness/PCBA Description'}
b_h_comp = b_h.rename(columns=b_h_rename)

#rename status to status update as both the sheets have the same column name
b_h_comp = b_h_comp.rename(columns={'Status': 'Status Update'})

#drop the Current QPV, Future QPV, QPV, Ship to Location columns
b_h_comp = b_h_comp.drop(columns=['Current QPV','Future QPV','QPV','Ship to Location'])
#print(b_h_comp.head())

#Load forecast into harnesses sheet
#print(comp.columns)
#print(b_h_comp.columns)
comp = pd.merge(comp,b_h_comp,how='left')
#print(comp.shape)
print(comp.head())

#Parse forecast qty columns
print(list(comp.columns))
comp_num_col_start = str(input("Please enter last column name before forecast data for components sheet: "))
comp_num_col_start_indx = comp.columns.get_loc(comp_num_col_start) + 1
comp_qtys = comp.iloc[:, comp_num_col_start_indx:]
print(comp_qtys.head())

#Multiply forecast qty by QPV
#comp['Qty'] = pd.to_numeric(comp['Qty'], downcast='float')
comp_forecast = comp_qtys.mul(comp['Qty'],axis=0)
#print(comp_forecast.head())

#Update board and harness forecast numbers to reflect multiplied QPV
comp.iloc[:, comp_num_col_start_indx:] = comp_forecast

#drop the status column
comp = comp.drop(columns=['Status'])
print(comp.head())
#print(comp.shape)
print("Done")

#Data Summary for each suppliers
print(comp['Manufacturer'].unique())
manufacturer = comp['Manufacturer'].unique()
print(manufacturer)
n = len(manufacturer)

#Change b_h sheet column names for comp sheet
b_h_rename = {'Description':'Harness/PCBA Description'}
b_h_comp = b_h.rename(columns=b_h_rename)

#rename status to status update as both the sheets have the same column name
b_h_comp = b_h_comp.rename(columns={'Status': 'Status Update'})

#drop the Current QPV, Future QPV, QPV, Ship to Location columns
b_h_comp = b_h_comp.drop(columns=['Current QPV','Future QPV','QPV','Ship to Location'])
#print(b_h_comp.head())

#Load forecast into harnesses sheet
#print(comp.columns)
#print(b_h_comp.columns)
comp = pd.merge(comp,b_h_comp,how='left')
#print(comp.shape)
print(comp.head())

#Parse forecast qty columns
print(list(comp.columns))
comp_num_col_start = str(input("Please enter last column name before forecast data for components sheet: "))
comp_num_col_start_indx = comp.columns.get_loc(comp_num_col_start) + 1
comp_qtys = comp.iloc[:, comp_num_col_start_indx:]
print(comp_qtys.head())

#Multiply forecast qty by QPV
#comp['Qty'] = pd.to_numeric(comp['Qty'], downcast='float')
comp_forecast = comp_qtys.mul(comp['Qty'],axis=0)
#print(comp_forecast.head())

#Update board and harness forecast numbers to reflect multiplied QPV
comp.iloc[:, comp_num_col_start_indx:] = comp_forecast

#drop the status column
comp = comp.drop(columns=['Status'])
print(comp.head())
#print(comp.shape)
print("Done")

#Data Summary for each suppliers
print(comp['Manufacturer'].unique())
manufacturer = comp['Manufacturer'].unique()
print(manufacturer)
n = len(manufacturer)
#Change b_h sheet column names for comp sheet
b_h_rename = {'Description':'Harness/PCBA Description'}
b_h_comp = b_h.rename(columns=b_h_rename)

#rename status to status update as both the sheets have the same column name
b_h_comp = b_h_comp.rename(columns={'Status': 'Status Update'})

#drop the Current QPV, Future QPV, QPV, Ship to Location columns
b_h_comp = b_h_comp.drop(columns=['Current QPV','Future QPV','QPV','Ship to Location'])
#print(b_h_comp.head())

#Load forecast into harnesses sheet
#print(comp.columns)
#print(b_h_comp.columns)
comp = pd.merge(comp,b_h_comp,how='left')
#print(comp.shape)
print(comp.head())

#Parse forecast qty columns
print(list(comp.columns))
comp_num_col_start = str(input("Please enter last column name before forecast data for components sheet: "))
comp_num_col_start_indx = comp.columns.get_loc(comp_num_col_start) + 1
comp_qtys = comp.iloc[:, comp_num_col_start_indx:]
print(comp_qtys.head())

#Multiply forecast qty by QPV
#comp['Qty'] = pd.to_numeric(comp['Qty'], downcast='float')
comp_forecast = comp_qtys.mul(comp['Qty'],axis=0)
#print(comp_forecast.head())

#Update board and harness forecast numbers to reflect multiplied QPV
comp.iloc[:, comp_num_col_start_indx:] = comp_forecast

#drop the status column
comp = comp.drop(columns=['Status'])
print(comp.head())
#print(comp.shape)
print("Done")

#Data Summary for each suppliers
print(comp['Manufacturer'].unique())
manufacturer = comp['Manufacturer'].unique()
print(manufacturer)
n = len(manufacturer)

#Write dataframes to Excel file and save to local drive
fp = str(input("Please enter output file name:\n"))
#savefp = input("Indicate filepath to save file in:\n")

#generating separate component level forecast sheet for each suppliers
with pd.ExcelWriter(fp + '.xlsx') as writer:
    tsla_forecast.to_excel(writer, index=False, sheet_name='High Level Forecast')
    b_h.to_excel(writer, index=False, sheet_name='Harnesses_Boards Forecast')
    comp.to_excel(writer, index=False, sheet_name='Component_Forecast')
    for i in range(0,n):
        comp[comp.Manufacturer == str(manufacturer[i])].to_excel(writer, index=False, sheet_name=str(manufacturer[i]))

for i in range(0,n):
    with pd.ExcelWriter(str(manufacturer[i]) + '.xlsx') as writer:
        comp[comp.Manufacturer == str(manufacturer[i])].to_excel(writer, index=False, sheet_name=str(manufacturer[i]))
        
    
print('Saved')
