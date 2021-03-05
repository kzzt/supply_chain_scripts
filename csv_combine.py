import pandas as pd
import glob
from datetime import date
import numpy as np

# pd.options.mode.chained_assignment = None

report_files = glob.glob(
    "C:\\Users\\u6zn\\Desktop\\writeoff\\returned\\*.xlsx", recursive=False,
)
# df = pd.concat((pd.read_excel(f, header=0) for f in report_files))
df = pd.concat((pd.read_excel(f, header=0) for f in report_files))

# print(df['STOREROOM'])

df.columns = map(str.upper, df.columns)

df["STOREROOM"] = df["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

df.drop_duplicates(keep="first", inplace=True, ignore_index=False)

df.to_excel("c:\\users\\u6zn\\desktop\\writeoff\\output_final.xlsx", index=False)

"""
# sorted_df = df.sort_values(
#     by=["STOREROOM", "ITEM #", "REPORT"], ascending=[True, True, False],
# )
sorted_df = df.sort_values(
    by=["STOREROOM", "REPORT", "ITEM #"], ascending=[True, False, True]
)
sorted_df.drop_duplicates(subset=["ITEM #", "STOREROOM"], keep="first", inplace=True)

sorted_df["STOREROOM"] = sorted_df["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

sorted_df = sorted_df.reset_index(drop=True)

print(sorted_df.dtypes)

# sorted_df = sorted_df.astype(
#     {
#         "ITEM #": "int",
#         "DESCRIPTION": "object",
#         "SITE": "object",
#         "STOREROOM": "object",
#         "REPORT": "object",
#         "STANDARD COST": "float64",
#         "ROP": "int64",
#         "MAX": "int64",
#         "ON HAND": "int64",
#         "ON-HAND VALUE": "float64",
#         "ON ORDER": "int64",
#         "IN TRANSIT": "int64",
#         "ON PR": "int64",
#         "SOFT RESERVED": "int64",
#         "HARD RESERVED": "int64",
#         "LAST ISSUE DATE": "datetime64",
#         "INVENTORY REVIEW DATE": "datetime64",
#         "CALCULATED EXCESS": "int64",
#         "EXTENDED CALCULATED EXCESS VALUE ": "int64",
#         "DECISION MAKER'S NAME": "object",
#         "DECISION": "object",
#         "REASON FOR RETENTION": "object",
#         "COMMENTS": "object",
#     }
# )

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
"""
