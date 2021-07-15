import pandas as pd
import glob

# report_files = glob.glob(
#     "C:\\Users\\u6zn\\Desktop\\writeoff\\returned\\*.xlsx", recursive=False,
# )

all_data = pd.DataFrame()
for f in report_files:
    xl = pd.ExcelFile(f)
    print(f, xl.sheet_names)
    df = pd.read_excel(f, header=[0,1], sheet_name="WOOD POLE VALIDATION")
    print(f, len(df.columns), df.head(1))

    # df.dropna(axis=0, subset=["ITEM #", "STOREROOM"], inplace=True)
    # df["STOREROOM"] = df["STOREROOM"].astype("str").replace("\.0", "", regex=True)
    # df["STOREROOM"] = df["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

    all_data = all_data.append(df)


# all_data.columns = map(str.upper, df.columns)

# all_data["STOREROOM"] = all_data["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

all_data.drop_duplicates(keep="first", inplace=True, ignore_index=True)

all_data.to_excel(
    "c:\\users\\u6zn\\desktop\\Poles\\output_poles_combined.xlsx", index=True
)
