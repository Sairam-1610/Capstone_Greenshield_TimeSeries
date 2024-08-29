#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Mar 22 08:04:20 2024

@author: sairamkumaran
"""

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

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


grouped_interval_merge_df = interval_merge_df.groupby(['year', 'month', 'day','ActivityType', 'EmployeeID'])[['ActivityHours', 'Transactions']].sum().reset_index()
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


from scipy.stats import zscore




######################################################################
# Convert ActivityHours to ActivityDays
merged_df['ActivityDays'] = merged_df['ActivityHours'] / 24

# Calculate TransactionsEfficiencyPerHour by dividing Transactions by ActivityDays
merged_df['TransactionsEfficiencyPerDay'] = merged_df['Transactions'] / merged_df['ActivityDays']
# RPI Calculation
def calculate_ActivityType_transactions_efficiency_rpi(df):
    # Group by department, year, month, and ActivityType
    grouped = df.groupby(['DeptID', 'year', 'month', 'ActivityType'])
    
    # Calculate mean TransactionsEfficiencyPerHour for each group
    mean_transactions_efficiency_per_hour = grouped['Transactions'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values_activitytype = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_transactions_efficiency_per_hour.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_activitytype = row['ActivityType']
        
        # Filter rows based on DeptID, year, month, and ActivityType, excluding the current row
        filtered_rows = mean_transactions_efficiency_per_hour[(mean_transactions_efficiency_per_hour['DeptID'] == current_dept_id) &
                                                               (mean_transactions_efficiency_per_hour['year'] == current_year) &
                                                               (mean_transactions_efficiency_per_hour['month'] == current_month) &
                                                               (mean_transactions_efficiency_per_hour['ActivityType'] != current_activitytype)]
        
        # Calculate the mean TransactionsEfficiencyPerHour for the filtered rows
        mean_transactions_efficiency_per_hour_exclude_current = filtered_rows['Transactions'].mean()
        
        # Append the mean value to the list
        mean_values_activitytype.append(mean_transactions_efficiency_per_hour_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_transactions_efficiency_per_hour['MeanTransactionsExcludeCurrent'] = mean_values_activitytype
    
    # Calculate Relative Performance Index (RPI) as the ratio of TransactionsEfficiencyPerHour to the mean TransactionsEfficiencyPerHour of other ActivityTypes
    mean_transactions_efficiency_per_hour['RPI'] = mean_transactions_efficiency_per_hour['Transactions'] / mean_transactions_efficiency_per_hour['MeanTransactionsExcludeCurrent']
    
    return mean_transactions_efficiency_per_hour

# Call the function with your DataFrame
relative_activitytype_transactions_efficiency_rpi = calculate_ActivityType_transactions_efficiency_rpi(merged_df)

# Rank ActivityTypes based on RPI for the year 2023
activitytyperank_transactions_efficiency_2023 = relative_activitytype_transactions_efficiency_rpi[relative_activitytype_transactions_efficiency_rpi['year'] == 2023].groupby('ActivityType')['RPI'].mean().sort_values(ascending=False).reset_index()

# Calculate mean RPI of ActivityTypes for each month
monthly_ActivityType_rpi_mean = relative_activitytype_transactions_efficiency_rpi.groupby(['year', 'month', 'ActivityType'])['RPI'].mean().reset_index()

# Calculate mean RPI of ActivityTypes for each department in each month
monthly_dept_ActivityType_rpi_mean = relative_activitytype_transactions_efficiency_rpi.groupby(['DeptID', 'year', 'month', 'ActivityType'])['RPI'].mean().reset_index()

def calculate_ActivityType_activityhours_rpi(df):
    # Group by department, year, month, and ActivityType
    grouped = df.groupby(['DeptID', 'year', 'month', 'ActivityType'])
    
    # Calculate mean ActivityHours for each department, year, month, and ActivityType
    mean_activity_hours = grouped['ActivityHours'].mean().reset_index()
    
    # Initialize a list to store mean values
    mean_values_activitytype = []
    
    # Iterate over each group in the DataFrame
    for index, row in mean_activity_hours.iterrows():
        current_dept_id = row['DeptID']
        current_year = row['year']
        current_month = row['month']
        current_activitytype = row['ActivityType']
        
        # Filter rows based on DeptID, year, month, and ActivityType, excluding the current row
        filtered_rows = mean_activity_hours[(mean_activity_hours['DeptID'] == current_dept_id) &
                                             (mean_activity_hours['year'] == current_year) &
                                             (mean_activity_hours['month'] == current_month) &
                                             (mean_activity_hours['ActivityType'] != current_activitytype)]
        
        # Calculate the mean ActivityHours for the filtered rows
        mean_activity_hours_exclude_current = filtered_rows['ActivityHours'].mean()
        
        # Append the mean value to the list
        mean_values_activitytype.append(mean_activity_hours_exclude_current)
    
    # Add the list of mean values as a new column to the DataFrame
    mean_activity_hours['MeanActivityHoursExcludeCurrent'] = mean_values_activitytype
    
    # Calculate Relative Performance Index (RPI) as the ratio of ActivityHours to the mean ActivityHours of other ActivityTypes
    mean_activity_hours['RPI'] = mean_activity_hours['ActivityHours'] / mean_activity_hours['MeanActivityHoursExcludeCurrent']
    
    return mean_activity_hours

# Call the function with your DataFrame
relative_activitytype_activityhours_rpi = calculate_ActivityType_activityhours_rpi(merged_df)

# Rank ActivityTypes based on RPI for the year 2023
activitytyperank_activityhours_2023 = relative_activitytype_activityhours_rpi[relative_activitytype_activityhours_rpi['year'] == 2023].groupby('ActivityType')['RPI'].mean().sort_values(ascending=False).reset_index()

# Calculate mean RPI of ActivityTypes for each month
monthly_ActivityType_rpi_mean = relative_activitytype_transactions_efficiency_rpi.groupby(['year', 'month', 'ActivityType'])['RPI'].mean().reset_index()

# Calculate mean RPI of ActivityTypes for each department in each month
monthly_dept_ActivityType_rpi_mean = relative_activitytype_transactions_efficiency_rpi.groupby(['DeptID', 'year', 'month', 'ActivityType'])['RPI'].mean().reset_index()



# Rank based on z-scores
def rank_performance(df, metric_col):
    # Calculate mean and standard deviation for the metric column within each group
    agg_df = df.groupby(['year', 'month', 'ActivityType'], as_index=False)[metric_col].agg(['mean', 'std'])
    
    # Merge aggregated stats back to the original DataFrame
    df = pd.merge(df, agg_df, on=['year', 'month', 'ActivityType'], suffixes=('', '_group'))
    
    # Calculate z-score for TransactionsEfficiencyPerHour
    df['z_score'] = zscore(df[metric_col], ddof=1)
    
    # Rank ActivityTypes based on z-score within each month and ActivityType subgroup
    df['Rank'] = df.groupby(['year', 'month', 'ActivityType'])['z_score'].rank(ascending=False)
    
    return df

# Example usage
df_with_ranks = rank_performance(merged_df, 'ActivityHours')

def calculate_mean_z_scores(df):
    # Group the DataFrame by year, month, ActivityType, and DeptID
    grouped_df = df.groupby(['year', 'month', 'ActivityType', 'DeptID'])
    
    # Calculate the mean z-score for each group
    agg_df = grouped_df['z_score'].mean().reset_index()
    
    return agg_df

# Example usage
mean_z_scores_df = calculate_mean_z_scores(df_with_ranks)


# Reset the index to ensure 'year', 'month', and 'ActivityType' are treated as regular columns
# Define a function to rank DeptID based on z_score within each group
# Reset the index to ensure 'year', 'month', and 'ActivityType' are treated as regular columns
# Reset the index to ensure 'year', 'month', and 'ActivityType' are treated as regular columns
mean_z_scores_df.reset_index(drop=True, inplace=True)

# Define a function to rank DeptID based on z_score within each group
def rank_dept(df_group):
    df_group['Rank'] = df_group['z_score'].rank(ascending=False)  # Rank in descending order
    return df_group  # Return the ranked DataFrame

# Group by 'year', 'month', and 'ActivityType'
grouped_df = mean_z_scores_df.groupby(['year', 'month', 'ActivityType'])

# Apply the rank_dept function to each group using transform
mean_z_scores_df_with_ranks = grouped_df.apply(rank_dept).reset_index(drop=True)

# Optionally, sort the DataFrame
mean_z_scores_df_with_ranks.sort_values(by=['year', 'month', 'ActivityType', 'Rank'], inplace=True)

# Reset the index without creating a new one
mean_z_scores_df_with_ranks.reset_index(drop=True, inplace=True)