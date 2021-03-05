import time
import pandas as pd
# import csv


todays_date = time.strftime('%m-%d-%Y')

print(todays_date)

colnames = ['date', 'dow', 'name']
data_file = pd.read_csv("c:\\users\\u6zn\\pythontasks\\dates.csv", names=colnames)

print(data_file.head())

# with open("c:\\users\\u6zn\\pythontasks\\dates.csv") as csv_file:
#     if todays_date in csv_file:
#         print(todays_date)
#     else:
#         print("not today")

dates = data_file.date.tolist()

if todays_date in dates:
    print(f"Today is {todays_date}. Go to work!")
else:
    print("Stay home today!!!")