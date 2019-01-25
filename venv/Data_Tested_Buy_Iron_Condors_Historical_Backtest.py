import pandas as pd
import xlsxwriter
import datetime
import os
import pickle

start = datetime.datetime.now()

class Week:
    def __init__(self, open, close, beginning_date, ending_date):
        self.open = open
        self.close = close
        self.beginning_date = beginning_date
        self.ending_date = ending_date
        self.put_option_chain = {'strike': [], 'premium': []} #stike, premium
        self.call_option_chain = {'strike': [], 'premium': []} #strike , premium

    def calc_profit(self):
        call_buy_index = 0
        while self.open > self.call_option_chain['strike'][call_buy_index]:
            call_buy_index += 1
        try:
            call_sell_index = self.call_option_chain['strike'].index(self.call_option_chain['strike'][call_buy_index] + testing_width)
        except ValueError:
            call_buy_index += 1
            call_sell_index = self.call_option_chain['strike'].index(self.call_option_chain['strike'][call_buy_index] + testing_width)

        self.call_premium = self.call_option_chain['premium'][call_buy_index] - self.call_option_chain['premium'][call_sell_index]
        while self.call_premium > testing_premium / 2.0:
            call_buy_index += 1
            call_sell_index += 1
            self.call_premium = self.call_option_chain['premium'][call_buy_index] - self.call_option_chain['premium'][call_sell_index]

        self.put_option_chain['strike'].reverse()
        self.put_option_chain['premium'].reverse()
        put_buy_index = 0
        while self.open < self.put_option_chain['strike'][put_buy_index]:
            put_buy_index += 1
        try:
            put_sell_index = self.put_option_chain['strike'].index(self.put_option_chain['strike'][put_buy_index] - testing_width)
        except ValueError:
            put_buy_index += 1
            put_sell_index = self.put_option_chain['strike'].index(self.put_option_chain['strike'][put_buy_index] - testing_width)

        self.put_premium = self.put_option_chain['premium'][put_buy_index] - self.put_option_chain['premium'][put_sell_index]
        while self.put_premium > testing_premium / 2.0:
            put_buy_index += 1
            put_sell_index += 1
            self.put_premium = self.put_option_chain['premium'][put_buy_index] - self.put_option_chain['premium'][put_sell_index]

        self.total_premium = self.call_premium + self.put_premium
        self.buy_put_strike = self.put_option_chain['strike'][put_buy_index]
        self.sell_put_strike = self.put_option_chain['strike'][put_sell_index]
        self.buy_call_strike = self.call_option_chain['strike'][call_buy_index]
        self.sell_call_strike = self.call_option_chain['strike'][call_sell_index]

        if self.close >= self.buy_put_strike and self.close <= self.buy_call_strike:
            self.profit = -self.total_premium  * 100 * num_contracts
        elif self.close > self.buy_call_strike and self.close < self.sell_call_strike:
            self.profit = ((self.close - self.buy_call_strike) - self.total_premium) * 100 * num_contracts
        elif self.close < self.buy_put_strike and self.close > self.sell_put_strike:
            self.profit = ((self.buy_put_strike - self.close) - self.total_premium) * 100 * num_contracts
        elif self.close >= self.sell_call_strike or self.close <= self.sell_put_strike:
            self.profit = (testing_width - self.total_premium) * 100 * num_contracts
        else:
            raise Exception("Something went horribly wrong...")



first_year = int(input("Enter the beginning year: "))
last_year = int(input("Enter the last year: "))
testing_years = range(first_year, last_year + 1)
testing_premium = float(input("Enter the testing premium: "))
testing_width = float(input("Enter the testing width: "))
num_contracts = int(input("Enter the number of contracts: "))

weeks_list = []
for file in os.listdir('data'):
    df = pd.read_csv(f'data/{file}')
    new_week = True
    for i in range(1, len(df.index)):
        if int(df['Date'][i].split('-')[0]) in testing_years:
            current_date = datetime.date(int(df['Date'][i].split('-')[0]), int(df['Date'][i].split('-')[1]),
                                         int(df['Date'][i].split('-')[2]))
            next_date = datetime.date(int(df['Date'][i + 1].split('-')[0]), int(df['Date'][i + 1].split('-')[1]),
                                         int(df['Date'][i + 1].split('-')[2]))
            if new_week == True:
                new_week = False
                first_open = float(df['Close'][i])
                beginning_date = current_date
            elif next_date - current_date >= datetime.timedelta(days=3):
                weeks_list.append(Week(first_open, float(df['Close'][i]), beginning_date, current_date))
                new_week = True

for week in weeks_list:
    found_block = False
    next_week = False
    print('creating option chains')
    for file in os.listdir(f'Option_Data/{week.beginning_date.year}'):
        option_df = pd.read_csv(f'Option_Data/{week.beginning_date.year}/{file}')
        for i in range(1, len(option_df.index)):
            df_date = datetime.date(int(option_df['date'][i].split('-')[0]), int(option_df['date'][i].split('-')[1]),
                                    int(option_df['date'][i].split('-')[2]))
            if df_date == week.beginning_date and option_df[' symbol'][i].split(' ')[0] == 'SPY':
                #adjustments may need to be made to statement after and depending on the format of the data.
                found_block = True
                print(df_date)
                exp_date = datetime.date(int(option_df[' expiration'][i].split('-')[0]),
                                         int(option_df[' expiration'][i].split('-')[1]),
                                         int(option_df[' expiration'][i].split('-')[2]))
                if abs(exp_date - week.ending_date) <= datetime.timedelta(days=1):
                    if option_df[' put/call'][i] == 'C':
                        week.call_option_chain['strike'].append(float(option_df[' strike'][i]))
                        week.call_option_chain['premium'].append(float(option_df[' price'][i]))
                    else:
                        week.put_option_chain['strike'].append(float(option_df[' strike'][i]))
                        week.put_option_chain['premium'].append(float(option_df[' price'][i]))
            elif found_block == True:
                next_week = True
                break
        if next_week == True:
            break

with open(f'pickle_data/{testing_years[0]}.pickle', 'wb') as pickle_out:
    pickle.dump(weeks_list, pickle_out)

# weeks_list = []
# for year in testing_years:
#     with open(f'pickle_data/{year}.pickle', 'rb') as pickle_in:
#         weeks_list.extend(pickle.load(pickle_in))

# for week in weeks_list:
#     if week.beginning_date == datetime.date(2012, 11, 5):
#         week_index = weeks_list.index(week)
#
# calls_df = pd.DataFrame(weeks_list[week_index].call_option_chain)
# puts_df = pd.DataFrame(weeks_list[week_index].put_option_chain)
# writer = pd.ExcelWriter('Debugging.xlsx', engine='xlsxwriter')
# calls_df.to_excel(writer, sheet_name='Calls')
# puts_df.to_excel(writer, sheet_name='puts')
# writer.save()


total_profit = 0
for week in weeks_list:
    week.calc_profit()
    if week.total_premium > .2 and week.profit <= (testing_width - week.total_premium) * 100 * num_contracts:
        total_profit += week.profit

finish = datetime.datetime.now()
print(f'\nCalculation finished in {finish - start}\n')

print('Report Summary:')
for week in weeks_list:
    print(f'Week beginning {week.beginning_date}')
    print(f'Profit: {week.profit}')
    print(f'Call Premium: {week.call_premium}')
    print(f'Put Premium: {week.put_premium}')
    print(f'Premium: {week.total_premium}')
    print(f'Open: {week.open}')
    print(f'Close: {week.close}')
    print(f'Buy Call Strike {week.buy_call_strike}')
    print(f'Buy Put Strike {week.buy_put_strike}\n')

print(f'${total_profit:,.2f}')