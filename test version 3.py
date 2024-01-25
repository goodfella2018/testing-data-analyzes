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

def transform_values(df, value):
    data = df
    grouped_data = data.groupby(data['Date'].dt.year)[value].mean()
    min_year = grouped_data.idxmin()
    max_year = grouped_data.idxmax()
    for year, mean in grouped_data.items():
        if year not in [min_year, max_year]:
            if mean < 0.5 * grouped_data.mean():
                grouped_data[year] = 0
    return df


def forecast_same_date_next_year(df, date_column, time_column, value_column):
    df['Datetime'] = pd.to_datetime(df[date_column].astype(str) + ' ' + df[time_column].astype(str))
    df = df.sort_values(by='Datetime')

    forecast_results = []

    for i in range(len(df)):
        current_row = df.iloc[i]
        actual_value = current_row[value_column]
        same_day_previous_years = df[(df['Datetime'].dt.month == current_row['Datetime'].month) &
                                     (df['Datetime'].dt.day == current_row['Datetime'].day) &
                                     (df['Datetime'].dt.hour == current_row['Datetime'].hour) &
                                     (df['Datetime'].dt.year < current_row['Datetime'].year)]

        # Average Forecast
        average_forecast = same_day_previous_years[value_column].mean()
        average_accuracy = 100 - np.abs(
            (actual_value - average_forecast) / actual_value * 100) if actual_value != 0 else None

        # SMA Forecast
        # window_sizes = [12, 24, 48]
        # sma_forecasts = {}
        # sma_accuracies = {}
        # for window_size in window_sizes:
        #     rolling_mean = same_day_previous_years[value_column].rolling(window=window_size).mean()
        #     if not rolling_mean.empty:
        #         sma_forecast = rolling_mean.iloc[-1]
        #         sma_accuracy = 100 - np.abs(
        #             (actual_value - sma_forecast) / actual_value * 100) if actual_value != 0 else None
        #     else:
        #         sma_forecast = None
        #         sma_accuracy = None
        #     sma_forecasts[f'SMA_{window_size}'] = sma_forecast
        #     sma_accuracies[f'SMA_{window_size}_Accuracy'] = sma_accuracy

        # Linear Regression Forecast
        if not same_day_previous_years.empty:

            years_of_data = same_day_previous_years['Datetime'].dt.year.nunique()
            points_per_year = max(24 // years_of_data, 1)

            regression_data = pd.DataFrame()
            for year in same_day_previous_years['Datetime'].dt.year.unique():
                year_data = same_day_previous_years[same_day_previous_years['Datetime'].dt.year == year]
                target_day_data = year_data[year_data['Datetime'] == current_row['Datetime']]
                regression_data = pd.concat([regression_data, target_day_data])

                extra_points = points_per_year - target_day_data.shape[0]
                if extra_points > 0:
                    additional_data = year_data.sample(min(extra_points, year_data.shape[0]), replace=False)
                    regression_data = pd.concat([regression_data, additional_data])

            regression_data = regression_data.head(24)
            X = np.array(list(range(len(regression_data)))).reshape(-1, 1)
            y = regression_data[value_column].values

            if len(y) > 0:
                model = LinearRegression().fit(X, y)
                forecast_lr = model.predict(np.array([[len(X)]]))[0]
                lr_accuracy = 100 - np.abs(
                    (actual_value - forecast_lr) / actual_value * 100) if actual_value != 0 else None
            else:
                forecast_lr = None
                lr_accuracy = None
        else:
            forecast_lr = None
            lr_accuracy = None

        forecast_results.append({
            'Datetime': current_row['Datetime'],
            'Actual Value': actual_value,
            'Average Forecast': average_forecast,
            'Average Accuracy': average_accuracy,
            # 'SMA_12 Forecast': sma_forecasts.get('SMA_12'),
            # 'SMA_12 Accuracy': sma_accuracies.get('SMA_12_Accuracy'),
            # 'SMA_24 Forecast': sma_forecasts.get('SMA_24'),
            # 'SMA_24 Accuracy': sma_accuracies.get('SMA_24_Accuracy'),
            # 'SMA_48 Forecast': sma_forecasts.get('SMA_48'),
            # 'SMA_48 Accuracy': sma_accuracies.get('SMA_48_Accuracy'),
            'Linear Regression Forecast': forecast_lr,
            'Linear Regression Accuracy': lr_accuracy
        })

    forecast_df = pd.DataFrame(forecast_results)
    forecast_df.sort_values('Datetime', inplace=True)

    return forecast_df

def forecast_next_day(df, date_column, time_column, value_column):
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

    forecast_results_short = forecast_next_day(df, 'Date', 'Time', third_column)
    forecast_result_long = forecast_same_date_next_year(df, 'Date', 'Time', third_column)
    with pd.ExcelWriter(output_path) as writer:
        min_max_years_table.to_excel(writer, sheet_name='MinMaxYears')
        last_three_years_table.to_excel(writer, sheet_name='LastThreeYears')
        forecast_results_short.to_excel(writer, sheet_name='Forecasting_short', index=False)
        forecast_result_long.to_excel(writer, sheet_name='Forecasting_long', index=False)
    subprocess.Popen([output_path], shell=True)

analyze_file('Foranalysis.xlsx', 'Dashbahesi')
