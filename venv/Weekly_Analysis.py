import pandas as pd
import xlsxwriter
import os
import datetime


class Year:
    def __init__(self, year):
        self.year = year
        self.weeks = []

    def calc_high_low_change(self):
        self.open = self.weeks[0].open
        self.close = self.weeks[-1].close
        self.percent_change = (self.close - self.open) / self.open

    def calc_week_performances(self):
        self.positive_weeks = 0
        self.negative_weeks = 0
        negative_changes = []
        positive_changes = []
        all_changes = []
        for week in self.weeks:
            if week.percent_change > 0:
                self.positive_weeks += 1
                positive_changes.append(week.percent_change)
                all_changes.append(week.percent_change)
            elif week.percent_change < 0:
                self.negative_weeks += 1
                negative_changes.append(week.percent_change)
                all_changes.append(week.percent_change)
        if len(negative_changes) == 0:
            negative_changes.append(0)
        self.average_week = sum(all_changes) / float(len(all_changes))
        self.average_positive_week = sum(positive_changes) / float(len(positive_changes))
        self.average_negative_week = sum(negative_changes) / float(len(negative_changes))
        self.worst_week = min(negative_changes)
        self.best_week = max(positive_changes)

    def calc_consequtive_negative_weeks(self):
        self.consequtive_negative_weeks = 0
        for i in range(len(self.weeks)):
            if i < len(self.weeks) - 1:
                if self.weeks[i].percent_change < 0 and self.weeks[i + 1].percent_change < 0:
                    self.consequtive_negative_weeks += 1


class Week:
    def __init__(self, date, open, close, volume):
        self.date = date
        self.open = open
        self.close = close
        self.volume = volume
        self.percent_change = (close - open) / open


years_dict = {}
for file in os.listdir('data'):
    df = pd.read_csv(f'data/{file}')
    for i in range(1, len(df.index)):
        raw_date_splits = df['Date'][i].split('-')
        date = datetime.date(int(raw_date_splits[0]), int(raw_date_splits[1]), int(raw_date_splits[2]))
        week = Week(date, float(df['Open'][i]), float(df['Close'][i]), float(df['Volume'][i]))
        if not week.date.year in years_dict:
            years_dict[week.date.year] = Year(week.date.year)
        years_dict[week.date.year].weeks.append(week)

total_positive_weeks = 0
total_negative_weeks = 0
for year in years_dict.values():
    year.calc_high_low_change()
    year.calc_week_performances()
    year.calc_consequtive_negative_weeks()
    total_positive_weeks += year.positive_weeks
    total_negative_weeks += year.negative_weeks

def sort_by_year(year):
    return year.year

year_list = sorted(years_dict.values(), key=sort_by_year)

out_dict = {'Year': [], 'Year Change': [], '+ Weeks': [], '- Weeks': [], 'Avg Week': [], 'Avg + Week': [],
            'Avg - Week': [], 'Best Week': [], 'Worst Week': [], 'Consequtive Negative Weeks': []}

for year in year_list:
    out_dict['Year'].append(year.year)
    out_dict['Year Change'].append(year.percent_change)
    out_dict['+ Weeks'].append(year.positive_weeks)
    out_dict['- Weeks'].append(year.negative_weeks)
    out_dict['Avg Week'].append(year.average_week)
    out_dict['Avg + Week'].append(year.average_positive_week)
    out_dict['Avg - Week'].append(year.average_negative_week)
    out_dict['Best Week'].append(year.best_week)
    out_dict['Worst Week'].append(year.worst_week)
    out_dict['Consequtive Negative Weeks'].append(year.consequtive_negative_weeks)

output_df = pd.DataFrame(out_dict)
writer = pd.ExcelWriter('output/output.xlsx', engine='xlsxwriter')
output_df.to_excel(writer, sheet_name='S&P 500')
writer.save()