# -*- coding: utf-8 -*-
"""
Created on Wed Feb 16 14:26:26 2022

@author: P0109482
"""
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 15 13:00:47 2022

@author: P0109482
"""
###############################
#### libraries ####
###############################
import pandas as pd
import numpy as np
import PySimpleGUI as sg
import time
import win32com.client as win32
from datetime import date

###############################
#### Main Function ####
###############################
def dailyEmailCleanData(df):
    ### Cleaning to be able to group by day and trend by day with only enter and exit dates
    #format to date_time
    df['Entry Date'] = pd.to_datetime(df['Entry Date'], errors = 'coerce')
    df['Exit Date']  = pd.to_datetime(df['Exit Date'], errors = 'coerce')
    file_date = str(date.today())
    file_date = pd.to_datetime(file_date).to_period('D')
    #Adding month columns and cartesian product of every day in the observed range
    all_date = df['Entry Date'].append(df['Exit Date']).unique()
    max_date = file_date.to_timestamp()  + pd.DateOffset(1)
    min_date = pd.to_datetime(np.nanmin(all_date))
    day_series = pd.date_range(start=min_date, end=max_date, freq='D') 
    day_series = pd.DataFrame(pd.DatetimeIndex(day_series).normalize())
    #concate room and NCS Site for unique room count
    df['NCS_Site_Name_room'] = df['NCS_Site_Name'] + '_' + str(df['Room#'])
    #select rows of interest
    df = df[['FirstName','LastName','Current_Status','County','NCS_Site_Name','NCS_Site_Name_room','# Clients','Entry Date','Exit Date']]
    cp_ncs = df.merge(day_series, how='cross')
    cp_ncs[0] = cp_ncs[0].rename('Date')
    cp_ncs.rename(columns={0:'Date'}, inplace=True)
    #Convert to year-month-day
    cp_ncs['Entry Date'] = pd.to_datetime(cp_ncs['Entry Date']).dt.to_period('D')
    cp_ncs['Exit Date'] = pd.to_datetime(cp_ncs['Exit Date']).dt.to_period('D')
    cp_ncs['Date'] = pd.to_datetime(cp_ncs['Date']).dt.to_period('D')
    #replace na date with business logic
    #If checked in and no exit date put exit date as the file dates date
    dt_bool = (cp_ncs['Current_Status'] == 'Checked In') & (pd.isna(cp_ncs['Exit Date']))
    cp_ncs.loc[dt_bool,['Exit Date']] = max_date.to_period('D')
    #if there is no entry date replace na with the exit date
    dt_bool = pd.isna(cp_ncs['Entry Date'])
    cp_ncs.loc[dt_bool,['Entry Date']] = cp_ncs.loc[dt_bool,['Exit Date']]
    #bool for what days client was checkied in
    cp_ncs['On This Day'] = (cp_ncs['Entry Date'] <= cp_ncs['Date']) & (cp_ncs['Exit Date'] >= cp_ncs['Date'])
    #group by county-day, filter out false day, sum Househols and client count
    cp_ncs = cp_ncs[cp_ncs['On This Day']]
    
    ### group by county and DAy- then calculate number of unique NCS sites
    pop_graph = cp_ncs.groupby(['County','Date']).agg({"# Clients": 'sum', "FirstName": 'nunique', "NCS_Site_Name": 'nunique', "NCS_Site_Name_room": 'nunique'})
    #Lag one day for change
    pop_graph = pop_graph.reset_index()
    pop_graph['Lag_Date'] = pop_graph['Date'] -1
    pop_lag = pop_graph.merge(pop_graph,left_on=['County','Date'],right_on=['County','Lag_Date'])
    pop_lag['Client Count Daily Change'] = pop_lag['# Clients_y'] - pop_lag['# Clients_x']
    #select only x columns
    pop_graph = pop_lag[['County','Date_x',"# Clients_x","FirstName_x","NCS_Site_Name_x","NCS_Site_Name_room_x",'Client Count Daily Change']]
    #Rename Columns ## Hotels	# Rooms	# People	Change
    pop_graph = pop_graph.rename(columns={'Date_x':'Date','# Clients_x':'Total Clients','FirstName_x':'Total House Holds','NCS_Site_Name_x': 'Hotel Count','NCS_Site_Name_room_x':'Room Count'})
    #Most recent date date
    Daily_email = pop_graph[pop_graph['Date'] == pop_graph['Date'].max()]
    Daily_email['Total Clients'] =  Daily_email['Total Clients'].astype(int)
    Daily_email['Client Count Daily Change'] =  Daily_email['Client Count Daily Change'].astype(int)
    return pop_graph, Daily_email

###############################
#### Read in Files ####
#### Browes for File names ###
###############################
sg.theme("DarkTeal2")
layout = [[sg.Text('Enter File Location')],
            [sg.Text('Master NCS List - 2021', size=(20, 1)), sg.InputText(key='ncs_fileName'), sg.FileBrowse()],
            [sg.Text('Master NCS List - 2022', size=(20, 1)), sg.InputText(key='ncs22_fileName'), sg.FileBrowse()],
            [sg.Text('Afgan Refugee', size=(20, 1)), sg.InputText(key='ref_fileName'), sg.FileBrowse()],
            [sg.Text('Email recipient List', size=(20, 1)), sg.InputText(key='email_recipient_filename'), sg.FileBrowse(target='email_recipient_filename')],
            [sg.Submit(), sg.Cancel()]]
### Building Window
window = sg.Window('My File Browser', layout, size=(650,250))
ncs_filename=''  
breakme = False 
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Submit" or event=="Cancel":
        ncs_filename = values["ncs_fileName"]
        ncs22_filename = values["ncs22_fileName"]
        ref_filename = values["ref_fileName"]
        email_recipient_filename = values["email_recipient_filename"]
        time.sleep(1)
        break
window.close()
## read in NCS ##
ncs = pd.read_excel(ncs_filename,sheet_name = 'Sheet1',header=2, engine = 'openpyxl')
ncs = ncs.dropna(axis=0,how='all')
ncs = ncs.dropna(axis=1,how='all')
##Read in NCS 22 ##
ncs22 = pd.read_excel(ncs22_filename,sheet_name = 'Sheet1',header=2, engine = 'openpyxl')
ncs22 = ncs22.dropna(axis=0,how='all')
ncs22 = ncs22.dropna(axis=1,how='all')
## Refugee data ##
#ref_filename = 'C:/Users/P0109482/OneDrive - Oregon DHSOHA/Documents/Jeff\'s Data Entry/Afghan Refugee Tracking.xlsx'
ref = pd.read_excel(ref_filename,sheet_name = 'Tracking',header=1, engine = 'openpyxl')
ref = ref.dropna(axis=0,how='all')
ref = ref.dropna(axis=1,how='all')
#select columns of interest and rename them
ref = ref[['County', 'Hotel', 'Room Number','Status', 'Last Name', 'First Name','# of Family Members','Unnamed: 12', 'Unnamed: 14']]
ref = ref.rename(columns={'First Name':'FirstName','Last Name':'LastName','Status':'Current_Status','County':'County','Hotel':'NCS_Site_Name',
                          'Room Number':'Room#','# of Family Members':'# Clients','Unnamed: 12':'Entry Date','Unnamed: 14':'Exit Date'})

###############################
#### Run function for each file ####
###############################
## run funcion for NCS ##
daily_email_ncs = dailyEmailCleanData(ncs)
pop_graph_ncs = daily_email_ncs[0]
Daily_email_ncs = daily_email_ncs[1]
## run funcion for NCS22 ##
daily_email_ncs22 = dailyEmailCleanData(ncs22)
pop_graph_ncs22 = daily_email_ncs22[0]
Daily_email_ncs22 = daily_email_ncs22[1]
## run funcion for Reffugee ##
daily_email_ref = dailyEmailCleanData(ref)
pop_graph_ref = daily_email_ref[0]
Daily_email_ref = daily_email_ref[1]


###############################
##### Send out Email ####
###############################
file_date = str(date.today())
#email list- if no list is provide than use default email
if len(email_recipient_filename)>1 : 
        email_recipients = pd.read_excel(email_recipient_filename,sheet_name = 'Sheet1',header=None, engine = 'openpyxl')
        email_recipients = ';'.join(email_recipients[0])
else : email_recipients = 'Noah.Robins@dhsoha.state.or.us'
#email list to ; list
#### Create email ####
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = email_recipients
mail.Subject = 'Daily Shelter Report'
body_text = 'Hey everyboady this is the daily report.'
#mail.HTMLBody = '<h2>HTML Message body</h2>' 
txt1 = '<html><p><h2>Daily Hotel (' + file_date +') Sheltering Totals - 2020 Wildfires:</h2> {v}<p/> '.format(v=Daily_email_ncs.to_html(index=False))
txt2 = '<p><h2>Daily Hotel (' + file_date +') Sheltering Totals - 2021 Wildfires:</h2> {v}<p/> '.format(v=Daily_email_ncs22.to_html(index=False))
txt3 = '<p><h2>Daily Hotel (' + file_date +') Sheltering Totals - Afgan Refugees:</h2>{v}<p/>  </html>'.format(v=Daily_email_ref.to_html(index=False))
#mail.HTMLBody = '<html><p>Daily Hotel (' + file_date +') Sheltering Totals - 2020 Wildfires:<p/> {0} <p>Daily Hotel (' + file_date +') Sheltering Totals - 2021 Wildfires:<p/> {1} <p>Daily Hotel (' + file_date +') Sheltering Totals - Afgan Refugees:<p/> {2} </html>'.format(Daily_email_ncs.to_html(index=False),Daily_email_ncs22.to_html(index=False),Daily_email_ref.to_html(index=False))  
mail.HTMLBody = txt1 + txt2 + txt3
mail.Send()
