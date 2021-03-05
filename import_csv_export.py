import pandas as pd


file = "c:\\users\\u6zn\\downloads\\Export.csv"

df = pd.read_csv(file, parse_dates=True)

df['Date'] = pd.to_datetime(df['Date'])

print(df.head())

df.to_csv("c:\\users\\u6zn\\desktop\\file_export.csv")