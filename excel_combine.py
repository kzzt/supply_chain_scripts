import pandas as pd
import glob

report_files = glob.glob(
    "C:\\Users\\u6zn\\Desktop\\writeoff\\returned\\*.xlsx", recursive=False,
)

all_data = pd.DataFrame()
for f in report_files:
    df = pd.read_excel(f)
    print(f, len(df.columns))
    df.dropna(axis=0, subset=["ITEM #", "STOREROOM"], inplace=True)
    df["STOREROOM"] = df["STOREROOM"].astype("str").replace("\.0", "", regex=True)
    df["STOREROOM"] = df["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

    all_data = all_data.append(df, ignore_index=True)


all_data.columns = map(str.upper, df.columns)

# all_data["STOREROOM"] = all_data["STOREROOM"].apply(lambda x: "{0:0>3}".format(x))

all_data.drop_duplicates(keep="first", inplace=True, ignore_index=False)

all_data.to_excel(
    "c:\\users\\u6zn\\desktop\\writeoff\\output_final_append_c.xlsx", index=False
)
