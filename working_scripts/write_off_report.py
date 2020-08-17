"""code to combine and format exported write off reports in xlsx format"""

import pandas as pd
import glob
from datetime import date
import numpy as np

pd.options.mode.chained_assignment = None

report_files = glob.glob("C:\\Users\\u6zn\\Desktop\\csv\\*.xlsx")
df = pd.concat((pd.read_excel(f, header=0).dropna(1) for f in report_files))

df["Decision Maker's Name"] = np.nan
df["Decision"] = np.nan
df["Reason For Retention"] = np.nan
df["Comments"] = np.nan

df.columns = map(str.upper, df.columns)

print(df.dtypes)

sorted_df = df.sort_values(
    by=["STOREROOM", "REPORT", "ITEM #"], ascending=[True, False, True]
)
sorted_df.drop_duplicates(subset=["ITEM #", "STOREROOM"], keep="first", inplace=True)

sorted_df["STOREROOM"] = sorted_df["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

sorted_df = sorted_df.reset_index(drop=True)

d = date.today()

sorted_df.rename(columns={"CALCULATED EXCESS": "REVIEW QUANTITY"}, inplace=True)
sorted_df.loc[sorted_df["REPORT"].ne("EXCESS"), "REVIEW QUANTITY"] = sorted_df[
    "ON HAND"
]
sorted_df["ON-HAND VALUE"] = sorted_df["ON HAND"] * sorted_df["STANDARD COST"]
sorted_df.loc[
    sorted_df["REPORT"].ne("EXCESS"), "EXTENDED CALCULATED EXCESS VALUE "
] = sorted_df["ON-HAND VALUE"]
sorted_df["INVENTORY REVIEW DATE"] = d.strftime("%m/%d/%Y")
sorted_df["LAST ISSUE DATE"].apply(lambda x: x.strftime("%m/%d/%Y"))

sorted_df["LAST ISSUE DATE"] = sorted_df["LAST ISSUE DATE"].dt.strftime("%m/%d/%Y")


final_df = sorted_df[
    [
        "ITEM #",
        "DESCRIPTION",
        "SITE",
        "STOREROOM",
        "REPORT",
        "STANDARD COST",
        "ROP",
        "MAX",
        "ON HAND",
        "ON-HAND VALUE",
        "ON ORDER",
        "IN TRANSIT",
        "ON PR",
        "SOFT RESERVED",
        "HARD RESERVED",
        "LAST ISSUE DATE",
        "INVENTORY REVIEW DATE",
        "REVIEW QUANTITY",
        "EXTENDED CALCULATED EXCESS VALUE ",
        "DECISION MAKER'S NAME",
        "DECISION",
        "REASON FOR RETENTION",
        "COMMENTS",
    ]
]

print(final_df.head())
print(final_df[final_df.index.duplicated()])

final_df.to_excel(
    f"2020 WRITE-OFF {d.strftime('%d-%m-%Y')}.xlsx",
    sheet_name="2020 WRITE-OFF",
    index=False,
)

# workbook = xlsxwriter.Workbook("c:\\users\\u6zn\\desktop\\csv\\write_off_2020.xlsx")
# worksheet = workbook.get_worksheet_by_name("excess")
#
# txt_a = "Decision Maker's Name"
#
# worksheet.write("b2", txt)
# worksheet.data_validation(
#     "B2", {"validate": "list", "source": ["open", "high", "close"]}
# )
#
#
# workbook.close()
