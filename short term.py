import pandas as pd
from scipy import stats
import numpy as np
import pygame
import tkinter as tk
from tkinter import filedialog
import os
import subprocess
from sklearn.linear_model import LinearRegression
from pandas.tseries.offsets import DateOffset

def previous_year_period(df, date_time_column, hours=30):
    previous_year = df[date_time_column] - DateOffset(years=1)
    start = previous_year - pd.Timedelta(hours=hours)
    end = previous_year + pd.Timedelta(hours=hours)
    return start, end

def transform_values(df, value):
    data = df
    grouped_data = data.groupby(data['Date'].dt.year)[value].mean()
    for year, mean in grouped_data.items():
        if mean < 0.3 * grouped_data.mean():
            grouped_data[year] = 0
    return df





def forecast_next_day(df, date_column, time_column, value_column)
    df = transform_values(df, value_column)
    df['Datetime'] = pd.to_datetime(df[date_column].astype(str) + ' ' + df[time_column].astype(str))

    df = df.sort_values(by='Datetime')


    forecast_results = []


    for i in range(1, len(df)):
        current_row = df.iloc[i]
        prev_row = df.iloc[i - 1]
        if i == 1:
            continue

        moving_avg = np.mean([current_row[value_column], prev_row[value_column]])

        same_day_previous_year = current_row['Datetime'] - DateOffset(years=1)
        prev_year_data = df[(df['Datetime'] == same_day_previous_year) & (df[value_column] != 0)]

        if prev_year_data.empty:
            forecast = moving_avg
        else:
            forecast = np.mean([moving_avg, prev_year_data[value_column].values[0]])

        if current_row[value_column] != 0:
            sma_accuracy = 100 - np.abs((current_row[value_column] - forecast) / current_row[value_column] * 100)
        else:
            sma_accuracy = None

        forecast_results.append({
            'Datetime': current_row['Datetime'],
            'Actual Value': current_row[value_column],
            'Forecast': forecast,
            'SMA Accuracy': sma_accuracy
        })

    forecast_df = pd.DataFrame(forecast_results)
    forecast_df.sort_values('Datetime', inplace=True)

    return forecast_df


def analyze_file(file_path, third_column):
    def generate_table(df, third_column, latest_year):
        table = df.groupby('Unique_ID')[third_column].agg(
            max='max',
            min='min',
            median='median',
            spread=lambda x: x.max() - x.min(),
            mean='mean'
        ).reset_index()

        latest_year_data = df[df['Year'] == latest_year]
        latest_year_mean = latest_year_data.groupby('Unique_ID')[third_column].mean().reset_index().rename(
            columns={third_column: 'Latest_Year_Mean'})
        table = pd.merge(table, latest_year_mean, on='Unique_ID', how='left')

        previous_year_data = df[(df['Year'] < latest_year) & (df['Year'] == latest_year - 1)]
        previous_year_mean = previous_year_data.groupby('Unique_ID')[third_column].mean().reset_index().rename(
            columns={third_column: 'Previous_Year_Mean'})
        table = pd.merge(table, previous_year_mean, on='Unique_ID', how='left')

        previous_years_data = df[df['Year'] < latest_year]
        previous_years_mean = previous_years_data.groupby('Unique_ID')[third_column].mean().reset_index().rename(
            columns={third_column: 'Previous_Years_Mean'})
        table = pd.merge(table, previous_years_mean, on='Unique_ID', how='left')

        table['Latest_Year_Mean'] = table['Latest_Year_Mean'].fillna(0)
        table['Previous_Year_Mean'] = table['Previous_Year_Mean'].fillna(0)
        table['Previous_Years_Mean'] = table['Previous_Years_Mean'].fillna(0)

        table['Previous_Year_Variation'] = ((table['Latest_Year_Mean'] - table['Previous_Year_Mean']) / table[
            'Latest_Year_Mean']) * 100
        table['Previous_Years_Variation'] = ((table['Latest_Year_Mean'] - table['Previous_Years_Mean']) / table[
            'Latest_Year_Mean']) * 100

        return table

    df = pd.read_excel(file_path)
    df.fillna(0, inplace=True)
    df[third_column] = pd.to_numeric(df[third_column], errors='coerce').fillna(0)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Unique_ID'] = df['Date'].dt.strftime('%m-%d') + ' ' + df['Time'].astype(str)
    df['Year'] = df['Date'].dt.year

    overall_latest_year = df['Year'].max()
    min_year = df['Year'].min()

    df_no_extreme_years = df[(df['Year'] != min_year) & (df['Year'] != overall_latest_year) & (df[third_column] != 0)]
    latest_year = df_no_extreme_years['Year'].max()

    min_max_years_table = generate_table(df_no_extreme_years, third_column, latest_year)

    df_last_three_years = df[df['Year'] >= (overall_latest_year - 2)]
    last_three_years_table = generate_table(df_last_three_years, third_column, overall_latest_year)

    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
    dir_path = os.path.dirname(os.path.realpath(file_path))

    output_path = os.path.join(dir_path, f"{third_column}_Enhanced_statistics.xlsx")

    forecast_results = forecast_next_day(df, 'Date', 'Time', third_column)

    with pd.ExcelWriter(output_path) as writer:
        min_max_years_table.rename(columns={'Previous_Year_Mean': 'Year_Before_Latest_Mean',
                                            'Previous_Years_Mean': 'All_Years_Before_Latest_Mean',
                                            'Previous_Year_Variation': 'Year_Before_Latest_Variation',
                                            'Previous_Years_Variation': 'All_Years_Before_Latest_Variation'}, inplace=True)
        last_three_years_table.rename(columns={'Previous_Year_Mean': 'Year_Before_Latest_Mean',
                                               'Previous_Year_Variation': 'Year_Before_Latest_Variation',
                                               'Previous_Years_Variation': 'All_Years_Before_Latest_Variation',
                                               'Previous_Years_Mean': 'All_Years_Before_Latest_Mean'}, inplace=True)
        min_max_years_table.to_excel(writer, sheet_name='MinMaxYears')
        last_three_years_table.to_excel(writer, sheet_name='LastThreeYears')
        forecast_results.to_excel(writer, sheet_name='Forecasting', index=False)
    subprocess.Popen([output_path], shell=True)

analyze_file('Foranalysis.xlsx', 'Dashbahesi')
