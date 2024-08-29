#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar 27 00:49:51 2024

@author: sairamkumaran
"""

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy.stats import zscore


# Load the Excel file
excel_file = "DeptartmentIntervalProductionResults.xlsx"

# Read the first sheet into a DataFrame
interval_employee_df = pd.read_excel(excel_file, sheet_name="IntervalEmployeeTimeTracking")

# Read the second sheet into a DataFrame
interval_completed_df = pd.read_excel(excel_file, sheet_name="IntervalCompletedItems")

interval_employee_df.columns = interval_employee_df.columns.str.replace(' ', '')

interval_completed_df.columns = interval_completed_df.columns.str.replace(' ', '')

# Extract year, month, and day from the "Date" column in Sheet1
interval_completed_df['year'] = interval_completed_df['DateofShiftStart'].dt.year
interval_completed_df['month'] = interval_completed_df['DateofShiftStart'].dt.month.astype(str).str.zfill(2)
interval_completed_df['day'] = interval_completed_df['DateofShiftStart'].dt.day.astype(str).str.zfill(2)



# Extract year, month, and day from the "Date" column in Sheet1
interval_employee_df['year'] = interval_employee_df['DateofShiftStart'].dt.year
interval_employee_df['month'] = interval_employee_df['DateofShiftStart'].dt.month.astype(str).str.zfill(2)
interval_employee_df['day'] = interval_employee_df['DateofShiftStart'].dt.day.astype(str).str.zfill(2)
interval_merge_columns = ['year', 'month', 'day', 'EmployeeID', 'GreenshieldIntervalOrder', 'IntervalID', 'DateRef']

# Merge the two selected dataframes on the specified columns

interval_merge_df = pd.merge(interval_employee_df, interval_completed_df[['year', 'month', 'day', 'EmployeeID','GreenshieldIntervalOrder','IntervalID', 'DateRef', 'Type','Code','Transactions']],
                     on=['year', 'month', 'day', 'EmployeeID','GreenshieldIntervalOrder','IntervalID', 'DateRef'], how='inner')

# Drop any duplicated columns resulting from the merge
interval_merge_df = interval_merge_df.loc[:, ~interval_merge_df.columns.duplicated()]
# Add a column 'ActivityHours' by dividing 'ActivitySeconds' by 360
interval_merge_df['ActivityHours'] = interval_merge_df['ActivitySeconds'] / 60
interval_merge_df['ActivityHours'] = interval_merge_df['ActivityHours'] / 60


grouped_interval_merge_df = interval_merge_df.groupby(['year', 'month', 'day','Code', 'EmployeeID'])[['ActivityHours', 'Transactions']].sum().reset_index()
# Add a column 'TransactionsPerHour' by dividing 'Transactions' by 'ActivityHours'
grouped_interval_merge_df['TransactionsPerHour'] = grouped_interval_merge_df['Transactions'] / grouped_interval_merge_df['ActivityHours']

merged_df = pd.DataFrame()

# Excel file and sheet names
excel_file_path = "/Users/sairamkumaran/GREENSHIELD/ReportforDeptsHoursSummary.xlsx"
sheet1_name = "DeptPayTypeBuckets"
sheet2_name = "DeptWorkingHoursByEmployee"

# Load data from Sheet2
df_DeptWorkingHoursByEmployee = pd.read_excel(excel_file_path, sheet_name=sheet2_name)

# Replace spaces in column names for Sheet2
df_DeptWorkingHoursByEmployee.columns = df_DeptWorkingHoursByEmployee.columns.str.replace(' ', '')

# Extract year, month, and day from the "Date" column in Sheet2
df_DeptWorkingHoursByEmployee['year'] = df_DeptWorkingHoursByEmployee['Date'].dt.year
df_DeptWorkingHoursByEmployee['month'] = df_DeptWorkingHoursByEmployee['Date'].dt.month.astype(str).str.zfill(2)
df_DeptWorkingHoursByEmployee['day'] = df_DeptWorkingHoursByEmployee['Date'].dt.day.astype(str).str.zfill(2)


merged_df = grouped_interval_merge_df.merge(df_DeptWorkingHoursByEmployee[['DeptID', 'EmployeeID', 'year', 'month', 'day']], 
                                        on=['EmployeeID', 'year', 'month', 'day'], 
                                        how='inner')
# Dropping any duplicated columns (if any)
merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]


merged_df['ActivityDays'] = merged_df['ActivityHours'] / 24

# Calculate TransactionsEfficiencyPerHour by dividing Transactions by ActivityDays
merged_df['TransactionsEfficiencyPerDay'] = merged_df['Transactions'] / merged_df['ActivityDays']

import pandas as pd

def calculate_Code_transactions_efficiency_rpi(df):
    # Group by department, year, month, and Code
    grouped = df.groupby(['DeptID', 'year', 'month', 'Code'])
    
    # Calculate mean TransactionsEfficiencyPerDay for each group
    mean_transactions_efficiency_per_day = grouped['TransactionsEfficiencyPerDay'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values_code = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_transactions_efficiency_per_day.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_code = row['Code']
        
        # Filter rows based on DeptID, year, month, and Code, excluding the current row
        filtered_rows = mean_transactions_efficiency_per_day[(mean_transactions_efficiency_per_day['DeptID'] == current_dept_id) &
                                                              (mean_transactions_efficiency_per_day['year'] == current_year) &
                                                              (mean_transactions_efficiency_per_day['month'] == current_month) &
                                                              (mean_transactions_efficiency_per_day['Code'] != current_code)]
        
        # Calculate the mean TransactionsEfficiencyPerDay for the filtered rows
        mean_transactions_efficiency_per_day_exclude_current = filtered_rows['TransactionsEfficiencyPerDay'].mean()
        
        # Append the mean value to the list
        mean_values_code.append(mean_transactions_efficiency_per_day_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_transactions_efficiency_per_day['MeanTransactionsEfficiencyPerDayExcludeCurrent'] = mean_values_code
    
    # Calculate Relative Performance Index (RPI) as the ratio of TransactionsEfficiencyPerDay to the mean TransactionsEfficiencyPerDay of other Codes
    mean_transactions_efficiency_per_day['RPI'] = mean_transactions_efficiency_per_day['TransactionsEfficiencyPerDay'] / mean_transactions_efficiency_per_day['MeanTransactionsEfficiencyPerDayExcludeCurrent']
    
    return mean_transactions_efficiency_per_day

# Call the function with your DataFrame
relative_code_transactions_efficiency_rpi = calculate_Code_transactions_efficiency_rpi(merged_df)

# Rank Codes based on RPI for the year 2023
coderank_transactions_efficiency_2023 = relative_code_transactions_efficiency_rpi[relative_code_transactions_efficiency_rpi['year'] == 2023].groupby('Code')['RPI'].mean().sort_values(ascending=False).reset_index()

# Calculate mean RPI of Codes for each month
monthly_Code_rpi_mean = relative_code_transactions_efficiency_rpi.groupby(['year', 'month', 'Code'])['RPI'].mean().reset_index()

# Calculate mean RPI of Codes for each department in each month
monthly_dept_Code_rpi_mean = relative_code_transactions_efficiency_rpi.groupby(['DeptID', 'year', 'month', 'Code'])['RPI'].mean().reset_index()

def calculate_ActivityHours_rpi(df):
    # Group by department, year, month, and code
    grouped = df.groupby(['DeptID', 'year', 'month', 'Code'])
    
    # Calculate mean ActivityHours for each department, year, month, and code
    mean_activity_hours = grouped['ActivityHours'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_activity_hours.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_code = row['Code']
        
        # Filter rows based on DeptID, year, month, and code, excluding the current row
        filtered_rows = mean_activity_hours[(mean_activity_hours['DeptID'] == current_dept_id) &
                                             (mean_activity_hours['year'] == current_year) &
                                             (mean_activity_hours['month'] == current_month) &
                                             (mean_activity_hours['Code'] != current_code)]
        
        # Calculate the mean ActivityHours for the filtered rows
        mean_activity_hours_exclude_current = filtered_rows['ActivityHours'].mean()
        
        # Append the mean value to the list
        mean_values.append(mean_activity_hours_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_activity_hours['MeanActivityHoursExcludeCurrent'] = mean_values
    
    # Calculate Relative Performance Index (RPI) as the ratio of ActivityHours to the mean ActivityHours of other codes
    mean_activity_hours['RPI'] = mean_activity_hours['ActivityHours'] / mean_activity_hours['MeanActivityHoursExcludeCurrent']
    
    return mean_activity_hours

# Call the function with your DataFrame
relative_code_activityhours_rpi = calculate_ActivityHours_rpi(merged_df)

# Rank codes based on RPI for the year 2023
coderank_2023 = relative_code_activityhours_rpi[relative_code_activityhours_rpi['year'] == 2023].groupby('Code')['RPI'].mean().sort_values(ascending=False).reset_index()

# Rank based on z-scores
def rank_performance(df, metric_col):
    # Calculate mean and standard deviation for the metric column within each group
    agg_df = df.groupby(['year', 'month', 'Code'], as_index=False)[metric_col].agg(['mean', 'std'])
    
    # Merge aggregated stats back to the original DataFrame
    df = pd.merge(df, agg_df, on=['year', 'month', 'Code'], suffixes=('', '_group'))
    
    # Calculate z-score for TransactionsEfficiencyPerDay
    df['z_score'] = zscore(df[metric_col], ddof=1)
    
    # Rank Codes based on z-score within each month and Code subgroup
    df['Rank'] = df.groupby(['year', 'month', 'Code'])['z_score'].rank(ascending=False)
    
    return df

# Example usage
df_with_ranks = rank_performance(merged_df, 'TransactionsEfficiencyPerDay')

def calculate_mean_z_scores(df):
    # Group the DataFrame by year, month, Code, and DeptID
    grouped_df = df.groupby(['year', 'month', 'Code', 'DeptID'])
    
    # Calculate the mean z-score for each group
    agg_df = grouped_df['z_score'].mean().reset_index()
    
    return agg_df

# Example usage
mean_z_scores_df = calculate_mean_z_scores(df_with_ranks)


# Reset the index to ensure 'year', 'month', and 'Code' are treated as regular columns
mean_z_scores_df.reset_index(drop=True, inplace=True)

# Define a function to rank DeptID based on z_score within each group
def rank_dept(df_group):
    df_group['Rank'] = df_group['z_score'].rank(ascending=False)  # Rank in descending order
    return df_group  # Return the ranked DataFrame

# Group by 'year', 'month', and 'Code'
grouped_df = mean_z_scores_df.groupby(['year', 'month', 'Code'])

# Apply the rank_dept function to each group using transform
mean_z_scores_df_with_ranks = grouped_df.apply(rank_dept).reset_index(drop=True)

# Optionally, sort the DataFrame
mean_z_scores_df_with_ranks.sort_values(by=['year', 'month', 'Code', 'Rank'], inplace=True)

# Reset the index without creating a new one
mean_z_scores_df_with_ranks.reset_index(drop=True, inplace=True)
