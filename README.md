trying to make automatic data analyzing tool, that takes xlsx data file as input and outputs new xlsx file with corresponding results (including various statistics and forecast).
as of right now the forecasting is done in following way:
1) short term - for short term we mean forecasting on the next day. For each day we take data on the exact time on exact year in last 2 days and also exact time on following 2 days in previous year and calculate their average
2) long term - there we have 2 methods, average and regression:
   for average we use same logic which is used in short term, the different thing is instead of previous 2 days and nexgt 2 days in previous year we take average of the exact time and date in all of the previous years
   for regression we require points to be exactly 24 (since for regression we need large data this is the only way to find out generally what will be the results in more complete data) if the data is not enough we randomly pick
   more then 1 points in each years data until the amount of points will be equal to 24, then we use basic regression of python.

