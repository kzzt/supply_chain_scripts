# Simplified program for calculating ROP from annual usage and lead time data

import time
import cx_Oracle
import numpy as np
import pandas as pd
from scipy.stats import norm
from sqlalchemy import create_engine

import onccfg as cfg

user = cfg.username
password = cfg.password
sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

cstr = f"oracle://{user}:{password}@{sid}"

engine = create_engine(
    cstr, convert_unicode=False, pool_recycle=20, pool_size=100, echo=False
)


def read_sql(filename):
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()

    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)

    return df_from_sql


timestring = time.strftime("%Y-%m-%d")

# for testing, using CSV instead of SQL
# df1 = pd.DataFrame(pd.read_csv("c:\\users\\u6zn\\desktop\\usage.csv"))
# df2 = pd.DataFrame(pd.read_csv("c:\\users\\u6zn\\desktop\\inv_param.csv"))
# df3 = pd.DataFrame(pd.read_csv("c:\\users\\u6zn\\desktop\\leadtime.csv"))
#

# df1 = read_sql(cfg.trinity_wkly)  # ITEMNUM: object datatype
df1 = read_sql(cfg.trinity_wkly)
df2 = read_sql(cfg.inventory_params)
df3 = read_sql(cfg.leadtime)
# df4 = abc_classification.abc_analysis(abc_classification.query, "TEST")


df1["ITEMNUM"] = df1["ITEMNUM"].astype(int)
df1["LOCATION"] = df1["LOCATION"].astype(str)
df1.sort_values(by=["LOCATION", "ITEMNUM", "WEEK"], inplace=True)

df2["LOCATION"] = df2["LOCATION"].astype(str)

# df4 = df4[["ITEMNUM", "ABC", "XYZ"]]


# reshape historical data -->
df1_pivot = (
    df1.groupby(["ITEMNUM", "LOCATION"])
    .agg(
        WEEKLY_AVG=pd.NamedAgg(column="NET_USAGE", aggfunc=np.mean),
        WEEKLY_STDDEV=pd.NamedAgg(column="NET_USAGE", aggfunc=np.std),
    )
    .reset_index()
)
print(df1_pivot.head())
merged = df1_pivot.merge(df2, how="left", on=["ITEMNUM", "LOCATION"])

merged = merged.merge(df3, how="left", on="ITEMNUM")

stdpk_regex = r"((?<!\S)(^\/)?(\b[0-9]+\s+\w{2}|[0-9]+)\/+([a-zA-Z]{2})\b)"
minlevel = merged["MINLEVEL"]

regex_df = merged["DESCRIPTION"].str.extractall(stdpk_regex)
regex_df = regex_df.groupby(level=0)[0].apply(list)
merged["STD_PKG"] = regex_df

merged.loc[merged["CALC_AVG_LT"].isnull(), "CALC_AVG_LT"] = merged["MAXIMO_LT"]
merged.loc[merged["CALC_STDDEV_LT"].isnull(), "CALC_STDDEV_LT"] = round(
    merged["MAXIMO_LT"] * 0.1, 0
)

avg_usage = merged["WEEKLY_AVG"] / 7
std_usage = merged["WEEKLY_STDDEV"] / 7
lead_time_avg = merged["CALC_AVG_LT"]
lead_time_std = merged["CALC_STDDEV_LT"]


service_level = 0.95

lead_time_demand = avg_usage * lead_time_avg
demand_variability = lead_time_avg * std_usage ** 2
lead_time_variability = (avg_usage * lead_time_std) ** 2
safety_stock = norm.ppf(service_level) * np.sqrt(
    demand_variability + lead_time_variability
)

merged["LT DEMAND"] = lead_time_demand
merged["DEMAND VARIABILITY"] = demand_variability
merged["LT VARIABILITY"] = lead_time_variability
merged["SAFETY STOCK"] = safety_stock
rop = round(safety_stock + lead_time_demand, 0)
merged["ROP"] = rop
try:
    merged.to_excel(
        f"c:\\users\\u6zn\\desktop\\ROP REVIEW {timestring}.xlsx", index=False
    )
except:
    merged.to_excel(
        f"c:\\users\\u6zn\\desktop\\ROP REVIEW {timestring} v2.xlsx", index=False
    )
