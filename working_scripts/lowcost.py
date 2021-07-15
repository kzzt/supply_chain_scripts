import pandas as pd


## get contact list
contact_list = pd.read_excel(
    "c:\\users\\u6zn\\desktop\\contact_list\\contact_list.xlsx"
)

print(contact_list.head())

## get lowcost report

## LOCATION = WHSE in contact list

## match location on lowcost report with whse in contact list
## send email to first name for each location/whse match
