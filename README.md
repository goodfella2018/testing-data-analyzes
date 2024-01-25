trying to make an automatic data analyzing tool, that takes xlsx data file as input and outputs a new xlsx file with corresponding results (including various statistics and forecasts).
as of right now, the forecasting is done in the following way:
1) Short term - for short term we mean forecasting on the next day. For each day we take data on the exact time of the exact year in the last 2 days and also the exact time on the following 2 days in the previous year and calculate their average
2) Long term - there we have 2 methods, average and regression:
   for average we use the same logic that is used in the short term, the difference is that instead of the previous 2 days and next 2 days in the previous year, we take an average of the exact time and date in all of the previous years
   for regression, we require points to be exactly 24 (since for regression we need large data this is the only way to find out generally what the results are in more complete data) if the data is not enough we randomly pick
   more than 1 point in each year's data until the number of points is equal to 24, then we use basic regression of Python.

