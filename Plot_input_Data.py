import win32com.client as win32
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

def reshape_dataframe1(df,column2_name):    
    #Five minute interval
    start_time = datetime.strptime('00:00', '%H:%M')
    
    reshaped_data = []
    for idx, row in df.iterrows():
        current_time = start_time
        for i in range(12):  # Columns
            reshaped_data.append([current_time.time(), row[1]])
            current_time += timedelta(minutes=5)  # Increase in five minute
        start_time += timedelta(hours=1)  # Increase in 1-hour
    
    reshaped_df = pd.DataFrame(reshaped_data, columns=['Time', column2_name])
    return reshaped_df

def reshape_dataframe2(df,column2_name):
    # Transform dataframe
    start_time = datetime.strptime('00:00', '%H:%M')
    
    reshaped_data = []
    for idx, row in df.iterrows():
        current_time = start_time
        for value in row[1:]:  # Columns
            reshaped_data.append([current_time.time(), value])
            current_time += timedelta(minutes=5)  # Increase in five minute
        start_time += timedelta(hours=1)  # Increase in 1-hour
    
    reshaped_df = pd.DataFrame(reshaped_data, columns=['Time', column2_name])
    return reshaped_df

df_total = pd.read_excel(os.getcwd()+"\\Data\\robust model_data.xlsx", sheet_name=["Price_DA", "Price_DARS", "Price_RT","Price_RTRS","Expected_P_RT_WPR"])

df_DA = df_total["Price_DA"].drop(labels=df_total["Price_DA"].columns[2],axis=1)
df_DA = df_DA.drop(labels=df_total["Price_DA"].columns[3],axis=1)
df_DA = df_DA.drop(labels=df_total["Price_DA"].columns[4],axis=1)
df_DA = df_DA.drop(labels=df_total["Price_DA"].columns[5],axis=1)
df_DA = reshape_dataframe1(df_DA,"Day-ahead price")
df_DA.set_index('Time', inplace = True)           # Index
#df_DA.plot()

df_DARS = df_total["Price_DARS"]
df_DARS = df_DARS.drop(labels=df_total["Price_DARS"].columns[2],axis=1)
df_DARS = df_DARS.drop(labels=df_total["Price_DARS"].columns[3],axis=1)
df_DARS = df_DARS.drop(labels=df_total["Price_DARS"].columns[4],axis=1)
df_DARS = df_DARS.drop(labels=df_total["Price_DARS"].columns[5],axis=1)
df_DARS = reshape_dataframe1(df_DARS,"Day-ahead reserve price")
df_DARS.set_index('Time', inplace = True)           # Index
#df_DARS.plot()

df_RT_temp = df_total["Price_RT"]
df_RT_temp = df_RT_temp.rename(columns = {'Real-time price' : 'Time'})
df_RT_temp = df_RT_temp.drop(labels=df_total["Price_RT"].columns[13],axis=1)
df_RT_temp = df_RT_temp.drop(labels=df_total["Price_RT"].columns[14],axis=1)
df_RT_temp = df_RT_temp.drop(labels=df_total["Price_RT"].columns[15],axis=1)
df_RT_temp = df_RT_temp.drop(labels=df_total["Price_RT"].columns[16],axis=1)


df_RT = reshape_dataframe2(df_RT_temp,"Real-time price")
df_RT.set_index('Time', inplace = True)           # Index
#df_RT.plot()

df_RTRS_temp = df_total["Price_RTRS"]
df_RTRS_temp = df_RTRS_temp.rename(columns = {'Real-time regulation price' : 'Time'})
df_RTRS_temp = df_RTRS_temp.drop(labels=df_total["Price_RTRS"].columns[13],axis=1)
df_RTRS_temp = df_RTRS_temp.drop(labels=df_total["Price_RTRS"].columns[14],axis=1)
df_RTRS_temp = df_RTRS_temp.drop(labels=df_total["Price_RTRS"].columns[15],axis=1)
df_RTRS= reshape_dataframe2(df_RTRS_temp,"Real-time reserve price")
df_RTRS.set_index('Time', inplace = True)           # Index
#df_RTRS.plot()

df = pd.concat([df_DA,df_DARS,df_RT,df_RTRS],axis=1)
plt.figure(figsize=(10, 6))

df.plot()

plt.xlabel('Time',fontweight= 'bold', size = 14)
plt.ylabel('Price ($/MW)',fontweight= 'bold', size = 14)
plt.title('Price data in day-ahead and real-time markets')
plt.legend()

plt.xticks(rotation=90)
selected_ticks = df.index[::12]
plt.xticks(selected_ticks)
plt.yticks(np.arange(0 , 60, 5))
plt.tight_layout()

#plt.show()

df_wind_temp = df_total["Expected_P_RT_WPR"]
df_wind_temp = df_wind_temp.rename(columns = {'Expected_P_RT_WPR' : 'Time'})
df_wind= reshape_dataframe2(df_wind_temp,"Expected wind power")
df_wind.set_index('Time', inplace = True)           # Index

plt.figure(figsize=(10, 6))

df_wind.plot()

plt.xlabel('Time',fontweight= 'bold', size = 14)
plt.ylabel('Power(MW)',fontweight= 'bold', size = 14)
plt.title('Expected wind power generation')
plt.legend()

plt.xticks(rotation=90)
selected_ticks = df_wind.index[::12]
plt.xticks(selected_ticks)
plt.yticks(np.arange(2 , 5, 0.5))
plt.tight_layout()

plt.show()
