# SPY_Iron_Condors
This program serves as a historical backtest to show to profitability of buying or selling weekly expiration iron condors for any given year or range of years.

The data folder just contains the daily data for the SPY ETF dating back to 2000. I downloaded the free csv file from Yahoo Finance.

You must purchase option data from Optionisitcs (http://www.optionistics.com/s/download_historical_data). That data goes inside the ‘Option_Data’ directory. Please place all csv files for a single year inside a folder entitled with the year of the data, and place that folder inside the ‘Option_Data’ directory.

The Data_Tested_Buy_Iron_Condors_Historical_Backtest.py script uses the actual option prices to give you an exact value for the profit/loss for a given max premium, width, and year range.

The Buy_Iron_Condors_Historical_Backtest.py and Sell_Iron_Condors_Historical_Backtest.py scripts give you an approximation of profit/loss based on the percent change between the current price and the first strike of each leg and the total premium paid for any given width and range of years.
