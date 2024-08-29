#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar 27 01:47:11 2024

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


grouped_interval_merge_df = interval_merge_df.groupby(['year', 'month', 'day','Type', 'EmployeeID'])[['ActivityHours', 'Transactions']].sum().reset_index()
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

# Convert ActivityHours to ActivityDays
merged_df['ActivityDays'] = merged_df['ActivityHours'] / 24
# Calculate TransactionsEfficiencyPerHour by dividing Transactions by ActivityDays
merged_df['TransactionsEfficiencyPerDay'] = merged_df['Transactions'] / merged_df['ActivityDays']

def calculate_Type_transactions_efficiency_rpi(df):
    # Group by department, year, month, and Type
    grouped = df.groupby(['DeptID', 'year', 'month', 'Type'])
    
    # Calculate mean TransactionsEfficiencyPerDay for each group
    mean_transactions_efficiency_per_day = grouped['Transactions'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values_type = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_transactions_efficiency_per_day.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_type = row['Type']
        
        # Filter rows based on DeptID, year, month, and Type, excluding the current row
        filtered_rows = mean_transactions_efficiency_per_day[(mean_transactions_efficiency_per_day['DeptID'] == current_dept_id) &
                                                              (mean_transactions_efficiency_per_day['year'] == current_year) &
                                                              (mean_transactions_efficiency_per_day['month'] == current_month) &
                                                              (mean_transactions_efficiency_per_day['Type'] != current_type)]
        
        # Calculate the mean TransactionsEfficiencyPerDay for the filtered rows
        mean_transactions_efficiency_per_day_exclude_current = filtered_rows['Transactions'].mean()
        
        # Append the mean value to the list
        mean_values_type.append(mean_transactions_efficiency_per_day_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_transactions_efficiency_per_day['MeanTransactionsPerDayExcludeCurrent'] = mean_values_type
    
    # Calculate Relative Performance Index (RPI) as the ratio of TransactionsEfficiencyPerDay to the mean TransactionsEfficiencyPerDay of other Types
    mean_transactions_efficiency_per_day['RPI'] = mean_transactions_efficiency_per_day['Transactions'] / mean_transactions_efficiency_per_day['MeanTransactionsPerDayExcludeCurrent']
    
    return mean_transactions_efficiency_per_day

# Call the function with your DataFrame
relative_type_transactions_efficiency_rpi = calculate_Type_transactions_efficiency_rpi(merged_df)

# Rank Types based on RPI for the year 2023
typerank_transactions_efficiency_2023 = relative_type_transactions_efficiency_rpi[relative_type_transactions_efficiency_rpi['year'] == 2023].groupby('Type')['RPI'].mean().sort_values(ascending=False).reset_index()

# Calculate mean RPI of Types for each month
monthly_Type_rpi_mean = relative_type_transactions_efficiency_rpi.groupby(['year', 'month', 'Type'])['RPI'].mean().reset_index()

# Calculate mean RPI of Types for each department in each month
monthly_dept_Type_rpi_mean = relative_type_transactions_efficiency_rpi.groupby(['DeptID', 'year', 'month', 'Type'])['RPI'].mean().reset_index()

def calculate_Type_activityhours_rpi(df):
    # Group by department, year, month, and type
    grouped = df.groupby(['DeptID', 'year', 'month', 'Type'])
    
    # Calculate mean ActivityHours for each department, year, month, and type
    mean_activity_hours = grouped['ActivityHours'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values_type = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_activity_hours.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_type = row['Type']
        
        # Filter rows based on DeptID, year, month, and type, excluding the current row
        filtered_rows = mean_activity_hours[(mean_activity_hours['DeptID'] == current_dept_id) &
                                             (mean_activity_hours['year'] == current_year) &
                                             (mean_activity_hours['month'] == current_month) &
                                             (mean_activity_hours['Type'] != current_type)]
        
        # Calculate the mean ActivityHours for the filtered rows
        mean_activity_hours_exclude_current = filtered_rows['ActivityHours'].mean()
        
        # Append the mean value to the list
        mean_values_type.append(mean_activity_hours_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_activity_hours['MeanActivityHoursExcludeCurrent'] = mean_values_type
    
    # Calculate Relative Performance Index (RPI) as the ratio of ActivityHours to the mean ActivityHours of other types
    mean_activity_hours['RPI'] = mean_activity_hours['ActivityHours'] / mean_activity_hours['MeanActivityHoursExcludeCurrent']
    
    return mean_activity_hours

# Call the function with your DataFrame
relative_type_activityhours_rpi = calculate_Type_activityhours_rpi(merged_df)

# Rank types based on RPI for the year 2023
typerank_activityhours_2023 = relative_type_activityhours_rpi[relative_type_activityhours_rpi['year'] == 2023].groupby('Type')['RPI'].mean().sort_values(ascending=False).reset_index()

def rank_performance(df, metric_col):
    # Calculate mean and standard deviation for the metric column within each group
    agg_df = df.groupby(['year', 'month', 'Type'], as_index=False)[metric_col].agg(['mean', 'std'])
    
    # Merge aggregated stats back to the original DataFrame
    df = pd.merge(df, agg_df, on=['year', 'month', 'Type'], suffixes=('', '_group'))
    
    # Calculate z-score for each performance metric value
    df['z_score'] = zscore(df[metric_col], ddof=1)
    
    # Rank the Department IDs based on z-score within each month and Type subgroup
    df['Rank'] = df.groupby(['year', 'month', 'Type'])['z_score'].rank(ascending=False)
    
    return df

# Example usage
df_with_ranks = rank_performance(merged_df, 'ActivityHours')

def calculate_mean_z_scores(df):
    # Group the DataFrame by year, month, Type, and DeptID
    grouped_df = df.groupby(['year', 'month', 'Type', 'DeptID'])
    
    # Calculate the mean z-score for each group
    agg_df = grouped_df['z_score'].mean().reset_index()
    
    return agg_df

# Example usage
mean_z_scores_df = calculate_mean_z_scores(df_with_ranks)

# Reset the index to ensure 'year', 'month', and 'Type' are treated as regular columns
mean_z_scores_df.reset_index(drop=True, inplace=True)

# Define a function to rank DeptID based on z_score within each group
def rank_dept(df_group):
    df_group['Rank'] = df_group['z_score'].rank(ascending=False)  # Rank in descending order
    return df_group  # Return the ranked DataFrame

# Group by 'year', 'month', and 'Type'
grouped_df = mean_z_scores_df.groupby(['year', 'month', 'Type'])

# Apply the rank_dept function to each group using transform
mean_z_scores_df_with_ranks = grouped_df.apply(rank_dept).reset_index(drop=True)

# Optionally, sort the DataFrame
mean_z_scores_df_with_ranks.sort_values(by=['year', 'month', 'Type', 'Rank'], inplace=True)

# Reset the index without creating a new one
mean_z_scores_df_with_ranks.reset_index(drop=True, inplace=True)
