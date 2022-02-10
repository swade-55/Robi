import streamlit as st
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestClassifier
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pandas.tseries.offsets import *
from openpyxl import load_workbook
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image
import time
import xlsxwriter
import openpyxl




st.write("""
# Robesonia Daily Production Report Builder App
This app produces daily report for C&S Robesonia facility.
""")

st.sidebar.header('User Input Features')

st.sidebar.markdown("""
[Example Triceps file](https://github.com/swade-55/Robi/blob/main/production_triceps_labor_report.xlsx?raw=true)
""")

# Collects user input features into dataframe
triceps_file = st.sidebar.file_uploader("Upload your input Triceps file", type=["xlsx"])

st.sidebar.markdown("""
[Example Qlik file](https://github.com/swade-55/Robi/blob/main/Robi%20Hours.xlsx?raw=true)
""")
qlik_file = st.sidebar.file_uploader("Upload your input Qlik file", type=["xlsx"])
check1 = st.sidebar.button("Analyze")


text_contents = '''
Foo, Bar
123, 456
789, 000
'''

if triceps_file is not None:
    df = pd.read_excel(triceps_file)
if qlik_file is not None:
    df9 = pd.read_excel(qlik_file)
if check1:
    df9 = df9.drop(columns = ['Warehouse','Week Ending','Shift','Status','FT/PT','Units','Indirect Hours','Productivity','Performance','Engagements','GER'
    ])
    df = df.drop(df.index[0])
    df = df.drop(df.index[0])
    df.columns = df.iloc[0]
    df = df[1:]

    def display(Triceps, Qlik):
        data = Triceps.copy()
        data['ACT_MINUTES'] = data['ACT_MINUTES'].astype(float, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(str, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(float, copy=False)
        data['EMPL_NUMBER'] = data['EMPL_NUMBER'].astype(int, copy=False)
        data['COMPLETED_CASES'] = data['COMPLETED_CASES'].astype(float, copy=False)
        data['IDLE_MIN'] = data['IDLE_MIN'].astype(float, copy=False)
        data['TASK'] = data['TASK'].astype(str, copy=False)
        data['START_DATE_TIME'] = data['START_DATE_TIME'].astype(str, copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        numbers = data.groupby(['Day', 'EMPL_NUMBER', 'WHSE'], as_index=False).sum()
        numbers['Performance'] = numbers['STD_MINUTES'] / (numbers['ACT_MINUTES'] + numbers['IDLE_MIN']) * 100
        numbers['Day'] = pd.to_datetime(numbers.Day)
        numbers['Day'] = numbers['Day'].dt.strftime('%m/%d/%Y')
        numbers['Day'] = pd.to_datetime(numbers.Day)
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        numbers['WHSE'] = numbers['WHSE'].astype(int)
        numbers['WHSE'] = numbers['WHSE'].map({1: 'GDC', 2: 'PDC', 3: 'FDC'})
        #numbers = numbers.drop(columns = ['Employee ID','Total Hours'])
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID', 'Date', 'Commodity'], right_on=['EMPL_NUMBER', 'Day', 'WHSE'])
        data['Uptime'] = (data['ACT_MINUTES'] + data['IDLE_MIN']) / (data['Total Hours'] * 60) * 100
        data['Date'] = data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data = data.rename(columns={'Commodity': 'Dept', 'Hire Date': 'DOH', 'Total Hours': 'Total_Hours'})
        data['DOH'] = data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data1 = data.groupby(['Employee ID', 'Name', 'Position', 'DOH', 'Dept', 'Supervisor'], as_index=False).sum()
        data1 = data1.drop(columns=['Performance', 'Uptime'])
        data1['Uptime'] = (data1['ACT_MINUTES']) / (data1['Total_Hours'] * 60) * 100
        data1['Performance'] = data1['STD_MINUTES'] / (data1['ACT_MINUTES'] + data1['IDLE_MIN']) * 100
        data1['Date'] = 'Total'
        data2 = data.append(data1)
        data2 = data2.drop(columns='IDLE_MIN')
        IDLE = data1[['Employee ID', 'IDLE_MIN']]
        data2 = data2.merge(IDLE, how='inner', left_on=['Employee ID'], right_on=['Employee ID'])
        Qlik = Qlik.rename(columns = {'Total Hours':'Total_Hours'})
        HoursWorked =  Qlik.groupby(['Employee ID'], as_index=False)['Total_Hours'].sum()
        HoursWorked = HoursWorked.rename(columns = {'Total_Hours':'Hours Worked'})
        data2 = data2.merge(HoursWorked, how='inner', left_on=['Employee ID'], right_on=['Employee ID'])
        data2['IDLE_MIN'] = data2['IDLE_MIN'] / 60
        data2 = data2.rename(columns={'IDLE_MIN': 'Idle Hours'})
        mypiv = data2.pivot(index=['Employee ID', 'Name', 'Position', 'DOH', 'Hours Worked', 'Idle Hours', 'Dept', 'Supervisor'],columns='Date')[['Performance', 'Uptime']].sort_values(by=['Dept', 'Position'], ascending=False)
        return mypiv

    def load(Triceps, Qlik):
        data = Triceps.copy()
        data['Pallets'] = 1
        data['ACT_MINUTES'] = data['ACT_MINUTES'].astype(float, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(str, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(float, copy=False)
        data['EMPL_NUMBER'] = data['EMPL_NUMBER'].astype(int, copy=False)
        data['COMPLETED_CASES'] = data['COMPLETED_CASES'].astype(float, copy=False)
        data['IDLE_MIN'] = data['IDLE_MIN'].astype(float, copy=False)
        data['TASK'] = data['TASK'].astype(str, copy=False)
        data['START_DATE_TIME'] = pd.to_datetime(data.START_DATE_TIME)
        data['START_DATE_TIME'] = pd.to_datetime(data.START_DATE_TIME) - timedelta(hours=5)
        data['START_DATE_TIME'] = data['START_DATE_TIME'].astype(str, copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        numbers = data.groupby(['Day', 'EMPL_NUMBER', 'WHSE'], as_index=False).sum()
        numbers['Day'] = pd.to_datetime(numbers.Day)
        numbers['Day'] = numbers['Day'].dt.strftime('%m/%d/%Y')
        numbers['Day'] = pd.to_datetime(numbers.Day)
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        numbers['WHSE'] = numbers['WHSE'].astype(int)
        numbers['WHSE'] = numbers['WHSE'].map({1: 'GDC', 2: 'PDC', 3: 'FDC'})
        #numbers = numbers.drop(columns = ['Employee ID','Total Hours'])
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID', 'Date'],right_on=['EMPL_NUMBER', 'Day'])
        data['Uptime'] = (data['ACT_MINUTES'] + data['IDLE_MIN']) / (data['Total Hours'] * 60) * 100
        data['Date'] = data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data['Scans/Hour'] = (data['Pallets'] / data['Total Hours'])
        data = data.rename(columns={'Commodity': 'Dept', 'Hire Date': 'DOH', 'Total Hours': 'Total_Hours'})
        data['DOH'] = data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data1 = data.groupby(['Employee ID', 'Name', 'Position', 'DOH', 'Dept', 'Supervisor'], as_index=False).sum()
        data1 = data1.drop(columns=['Scans/Hour', 'Uptime'])
        data1['Uptime'] = (data1['ACT_MINUTES'] + data1['IDLE_MIN']) / (data1['Total_Hours'] * 60) * 100
        data1['Scans/Hour'] = (data1['Pallets'] / data1['Total_Hours'])
        data1['Date'] = 'Total'
        data2 = data.append(data1)
        data2 = data2.drop(columns='IDLE_MIN')
        IDLE = data1[['Employee ID', 'IDLE_MIN']]
        data2 = data2.merge(IDLE, how='inner', left_on=['Employee ID'], right_on=['Employee ID'])
        data2['Hours Worked'] = (data2.groupby(['Employee ID', 'IDLE_MIN']).Total_Hours.transform('sum')) / 2
        data2['IDLE_MIN'] = data2['IDLE_MIN'] / 60
        data2 = data2.rename(columns={'IDLE_MIN': 'Idle Hours'})
        mypiv = data2.pivot_table(index=['Employee ID', 'Name', 'Position', 'DOH', 'Hours Worked', 'Idle Hours', 'Dept', 'Supervisor'],columns='Date',aggfunc='first')[['Scans/Hour', 'Uptime']].sort_values(by=['Dept', 'Position'], ascending=False)
        return mypiv

    def fork(Triceps, Qlik):
        data = Triceps.copy()
        data['ACT_MINUTES']= data['ACT_MINUTES'].astype(float,copy=False)
        data['STD_MINUTES']= data['STD_MINUTES'].astype(str,copy=False)
        data['STD_MINUTES']= data['STD_MINUTES'].astype(float,copy=False)
        data['EMPL_NUMBER']= data['EMPL_NUMBER'].astype(int,copy=False)
        data['COMPLETED_CASES']= data['COMPLETED_CASES'].astype(float,copy=False)
        data['IDLE_MIN']= data['IDLE_MIN'].astype(float,copy=False)
        data['TASK']= data['TASK'].astype(str,copy=False)
        data['START_DATE_TIME']= data['START_DATE_TIME'].astype(str,copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        data['Pallets'] = 1
        numbers = data.groupby(['Day','EMPL_NUMBER'],as_index=False).sum()
        numbers['Day'] = pd.to_datetime(numbers.Day)
        numbers['Day'] =numbers['Day'].dt.strftime('%m/%d/%Y')
        numbers['Day'] = pd.to_datetime(numbers.Day)
        #numbers = numbers.drop(columns = ['Employee ID','Total Hours'])
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID','Date'], right_on=['EMPL_NUMBER','Day'])
        data['Date'] =data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data = data.rename(columns = {'Hire Date':'DOH'})
        data['DOH'] = pd.to_datetime(data.DOH)
        data['DOH'] =data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data['Pallets/Hour'] = data['Pallets']/(data['ACT_MINUTES']/60)
        return data

    Puttriceps = df[df['JOB_CODE']=='PUT']
    Selecttriceps1 = df[df['JOB_CODE']=='CSL']
    Selecttriceps2 = df[df['JOB_CODE']=='CSE']
    Selecttriceps = Selecttriceps1.append(Selecttriceps2)
    Loadtriceps = df[df['JOB_CODE']=='LOD']
    Lettriceps = df[df['JOB_CODE']=='LET']

    forkhour3 = df9[df9['Position']=='Operator, Forklift - Step']
    forkhour1 = df9[df9['Position']=='Operator, Forklift']
    forkhour2 = df9[df9['Position']=='Forklift, Hourly, Freezer - Step']
    ForkQlik = forkhour1.append([forkhour3,forkhour2])
    selecthour3 = df9[df9['Position']=='Selector, Incentive, Freezer - Step']
    selecthour1 = df9[df9['Position']=='Selector, In Training']
    selecthour2 = df9[df9['Position']=='Selector, Incentive']       
    selecthour4 = df9[df9['Position']=='Selector, Incentive (ITT)']               
    SelectQlik = selecthour1.append([selecthour3,selecthour2,selecthour4])
    LoadQlik = df9[df9['Position']=='Loader - Step']

    Let = fork(Lettriceps, ForkQlik)
    Put = fork(Puttriceps,ForkQlik)
    selectors = display(Selecttriceps, SelectQlik)
    loaders = load(Loadtriceps,LoadQlik)

    Travel = df[df['JOB_CODE']=='TRV']
    Travel['START_DATE_TIME']= Travel['START_DATE_TIME'].astype(str,copy=False)
    Travel['Day'] = Travel['START_DATE_TIME'].str[0:10]
    Travel['Day'] = pd.to_datetime(Travel.Day)
    Travel['Day'] =Travel['Day'].dt.strftime('%m/%d/%Y')
    Travel['Day'] = pd.to_datetime(Travel.Day)
    Travel = Travel.groupby(['Day','EMPL_NUMBER'],as_index=False).sum()
    Travel = Travel.rename(columns = {'ACT_MINUTES':'TRV_MINUTES'})
    Travel = Travel.drop(columns = ['JOB_CODE','FACILITY','START_DATE_TIME','STD_MINUTES','IDLE_MIN','DELAY_MINUTES','COMPLETED_CUBE','COMPLETED_CASES','TASK'])
    Travel['EMPL_NUMBER'] = Travel['EMPL_NUMBER'].astype(int)

    ForkQlik = ForkQlik.rename(columns = {'Total Hours':'Total_Hours'})
    ForkQlik['Hours Worked'] = (ForkQlik.groupby('Employee ID').Total_Hours.transform('sum'))
    Let = fork(Lettriceps, ForkQlik)
    Let['Class'] = 'Letdowns/Hour'
    Let = Let.rename(columns = {'ACT_MINUTES':'Total_MINUTES'})
    Let = Let.rename(columns = {'Total Hours':'Total_Hours'})
    Put = fork(Puttriceps,ForkQlik)
    Put = Put.merge(Travel, how='left', left_on=['EMPL_NUMBER','Day'], right_on=['EMPL_NUMBER','Day'])
    Put = Put.drop(columns = ['Pallets/Hour'])
    Put['Total_MINUTES'] = Put['ACT_MINUTES']+Put['TRV_MINUTES']
    Put['Pallets/Hour'] = Put['Pallets']/(Put['Total_MINUTES']/60)
    Put['Class'] = 'Putaways/Hour'
    Put = Put.rename(columns = {'Total Hours':'Total_Hours'})
    Forks = Put.append(Let)
    Forks = Forks.rename(columns = {'Total Hours':'Total_Hours','Pallets/Hour':'Fork Metrics'})
    Forks2 = Forks.copy()
    Forks2 = Forks2.drop(columns = ['Class','Fork Metrics'])
    Forks2  = Forks2.groupby(['Employee ID','Name','Position','DOH','Commodity','Total_Hours','Hours Worked','Supervisor','Date'],as_index=False).sum()
    Forks2['Fork Metrics'] = (Forks2['Total_MINUTES']+Forks2['IDLE_MIN'])/(Forks2['Total_Hours']*60)*100
    Forks2['Class'] = 'Uptime'
    Forks3 = Forks.append(Forks2)

    Puttotal = Put.copy()
    Lettotal = Let.copy()
    Puttotal = Puttotal.rename(columns = {'Total Hours':'Total_Hours'})
    ptotal = Puttotal.groupby(['Employee ID','Name','Position','DOH','Commodity','Hours Worked','Supervisor'],as_index=False).sum()
    ptotal['Class'] = 'Putaways/Hour'
    ptotal = ptotal.drop(columns = 'Pallets/Hour')
    ptotal['Date'] = 'Total'
    ptotal['Pallets/Hour'] = ptotal['Pallets']/(ptotal['Total_MINUTES']/60)
    Lettotal = Lettotal.rename(columns = {'Total Hours':'Total_Hours'})
    ltotal = Lettotal.groupby(['Employee ID','Name','Position','DOH','Commodity','Hours Worked','Supervisor'],as_index=False).sum()
    ltotal['Class'] = 'Letdowns/Hour'
    ltotal['Date'] = 'Total'
    ltotal = ltotal.drop(columns = 'Pallets/Hour')
    ltotal['Pallets/Hour'] = ltotal['Pallets']/(ltotal['Total_MINUTES']/60)
    total = ptotal.append(ltotal)

    uptimetotal = total.copy()
    uptimetotal = uptimetotal.groupby(['Employee ID','Name','Position','DOH','Commodity','Hours Worked','Supervisor'],as_index=False).sum()
    uptimetotal = uptimetotal.drop(columns = 'Pallets/Hour')
    uptimetotal['Pallets/Hour'] = (uptimetotal['Total_MINUTES']+uptimetotal['IDLE_MIN'])/(uptimetotal['Hours Worked']*60)*100
    uptimetotal['Date'] = 'Total'
    uptimetotal['Class'] = 'Uptime'
    total = ptotal.append([ltotal,uptimetotal])
    total = total.rename(columns = {'Pallets/Hour':'Fork Metrics'})
    Forks4 = Forks3.append(total)
    Forks4 = Forks4.rename(columns = {'Commodity':'Dept'})
    Forks4 = Forks4.drop(columns = 'IDLE_MIN')
    IDLE = uptimetotal[['Employee ID','IDLE_MIN']]
    Forks4 = Forks4.merge(IDLE, how = 'inner', left_on = ['Employee ID'], right_on = ['Employee ID'])
    Forks4['Idle Hours'] = Forks4['IDLE_MIN']/60
    forkpiv = Forks4.pivot(index=['Employee ID','Name','Position','DOH','Hours Worked','Idle Hours','Dept','Supervisor'],columns=['Date','Class'],values = ['Fork Metrics']).sort_index(1)
    # Function to save all dataframes to one single excel
    def to_excel(df,df1,df2):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=True, sheet_name='Sheet1')
        df1.to_excel(writer, index=True, sheet_name='Sheet2')
        df2.to_excel(writer, index=True, sheet_name='Sheet3')
        workbook = writer.book
        #worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        #worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    df_xlsx = to_excel(forkpiv,selectors,loaders)
    st.download_button(label='📥 Download Current Result', data=df_xlsx ,file_name= 'df_test.xlsx')

        

    st.subheader('Schedule')
    st.write(forkpiv)

    st.subheader('Schedule')
    st.write(selectors)

    st.subheader('Schedule')
    st.write(loaders)



