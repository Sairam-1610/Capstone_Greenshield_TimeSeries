#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Mar 31 04:58:59 2024

@author: sairamkumaran
"""

import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


# Load DailySummaries dataset from Excel
gs_info = pd.read_excel("DailySummaries.xlsx", engine='openpyxl',index_col = 'Date' , parse_dates = True)
# Convert it to a DatetimeIndex
gs_info.index = pd.to_datetime(gs_info.index, unit='s')

#Replace space in column names
gs_info.columns = gs_info.columns.str.replace(' ','')
# Create a new column 'Weekend' and assign 1 if it's a weekend, otherwise 0
gs_info['Weekend'] = (gs_info.index.weekday >= 5).astype(int)

# Create a new column 'Weekday' and assign 1 if it's a weekday, otherwise 0
gs_info['Weekday'] = (gs_info.index.weekday < 5).astype(int)


# Selecting the desired columns
selected_columns = ['Dept04Processed','Dept04TotalInventoryBreakOut2',
                    'Dept11TotalCalls','Dept04TotalInventoryBreakOut3',
                    'Dep28Inventory','Weekday','Weekend','Dept05TotalCalls']

filtered_gs_info = gs_info.loc['2019':, selected_columns]


# Calculate correlation coefficients with 'Dept05TotalCalls'
correlations = filtered_gs_info.corrwith(filtered_gs_info['Dept05TotalCalls']).sort_values(ascending=False)

# Filter correlations between 0.4 and -0.2
filtered_correlations = correlations[(correlations > 0.4) | (correlations < -0.2)]

# Plot bar graph for the filtered columns
plt.figure(figsize=(12, 6))
bars = plt.bar(filtered_correlations.index, filtered_correlations, color=(filtered_correlations > 0).map({True: 'g', False: 'r'}))

# Add value labels on top of bars
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval + 0.01, f'{yval:.2f}', ha='center', va='bottom')

plt.title('Correlation of Columns with Dept05TotalCalls (from 2021 onwards)')
plt.xlabel('Columns')
plt.ylabel('Correlation Coefficient')
plt.xticks(rotation=45, ha='right')
plt.grid(axis='y', linestyle='--', alpha=0.7)

plt.show()

filtered_gs_info = filtered_gs_info.fillna(0)

train = filtered_gs_info.iloc[:1420,7]
test = filtered_gs_info.iloc[1420:,7]

exo = filtered_gs_info.iloc[:, :7]

exo_train = exo.iloc[:1420]
exo_test = exo.iloc[1420:]

#from statsmodels.tsa.seasonal import seasonal_decompose
#Decomp_single = seasonal_decompose(filtered_gs_info['Dept05TotalCalls'])
#Decomp_single.plot()
#######################IMPRTANT############################
from pmdarima import auto_arima
auto_arima(filtered_gs_info['Dept05TotalCalls'], exogenous = exo , m=6, trace = True, D=1).summary()
###########################################################
from statsmodels.tsa.statespace.sarimax import SARIMAX
Model = SARIMAX(train,exog= exo_train,order= (4,0,0),seasonal_order=(2,1,1,6))

Model =  Model.fit()

pprediction = Model.predict(len(train),len(train)+len(test)-1,exog=exo_test , typ='Levels')
# Apply condition to set negative values to zero
pprediction[pprediction < 0] = 0
weekends = pprediction.index[pprediction.index.weekday >= 5]
pprediction.loc[weekends] = 0
plt.plot(test,color = 'green' , label = 'Actual Total Calls')
plt.plot(pprediction,color = 'blue' , label = 'Predicted Total Calls')
plt.title("SARIMAX Model Predicting Dept05TotalCalls")
plt.xlabel("Days")
plt.ylabel("Total Calls")
plt.legend()
plt.show()

import math
from sklearn.metrics import mean_squared_error
rmse = math.sqrt(mean_squared_error(test,pprediction))

import pandas as pd
from statsmodels.tsa.stattools import grangercausalitytests

# Sample data (replace this with your actual data)
# Assuming gs_info is your DataFrame containing the relevant columns
selected_columns = ['Dept04Processed','Dept30Received','Dept04TotalInventoryBreakOut2',
                    'Dept11TotalCalls','Dept04TotalInventoryBreakOut3','Dept04TotalInventoryBreakOut1',
                    'Weekday','Dept04NewReceived','Dep28Inventory','Weekend','Dept05TotalCalls']
filtered_gs_info = gs_info[selected_columns]
filtered_gs_info = filtered_gs_info.fillna(0)
# Perform Granger causality test
def granger_causality_test(data, target_variable, max_lag=2):
    results = {}
    for column in data.columns:
        if column != target_variable:
            try:
                test_result = grangercausalitytests(data[[target_variable, column]], max_lag, verbose=False)
                p_values = [round(test_result[i+1][0]['ssr_ftest'][1], 4) for i in range(max_lag)]
                results[column] = p_values
            except Exception as e:
                print(f"An error occurred while performing Granger causality test for {target_variable} and {column}: {e}")
    return results

# Perform Granger causality test for each column against Dept05TotalCalls
granger_results = granger_causality_test(filtered_gs_info, 'Dept05TotalCalls')

# Display results
for column, p_values in granger_results.items():
    print(f"Granger causality test results for {column}:")
    for lag, p_value in enumerate(p_values, 1):
        print(f"Lag {lag}: p-value = {p_value}")

