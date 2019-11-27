import time

import cx_Oracle
import numpy as np
import pandas as pd
from scipy.stats import norm
from sqlalchemy import create_engine

import onccfg as cfg

timestring = time.strftime("%Y-%m-%d")


def oracle_conn():
    user = cfg.username
    password = cfg.password
    sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

    cstr = "oracle://{user}:{password}@{sid}".format(user=user, password=password, sid=sid)

    engine = create_engine(cstr, convert_unicode=False, pool_recycle=20, pool_size=100, echo=False)
    return engine


def read_sql(filename):
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()

    engine = oracle_conn()

    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)

    # df_from_sql.to_csv(f"c:\\users\\uxkp\\desktop\\{name_string}.csv", index=False)

    return df_from_sql


def abc_classification(perc):
    if 0 < perc < 0.8:
        return "A"
    elif 0.8 <= perc < 0.9:
        return "B"
    elif perc >= 0.9:
        return "C"


def xyz_classification(cv):
    if 0 < cv < 0.5:
        return "X"
    elif 0.5 <= cv < 1.5:
        return "Y"
    elif cv >= 1.5:
        return "Z"


query = read_sql("C:\\Users\\uxkp\\sql_queries\\scratch\\ABC_CLASS_PARAMETERS.sql")


def abc_analysis(data, type="Test"):
    data["ITEMNUM"] = data["ITEMNUM"].astype(int)
    data["LOCATION"] = data["LOCATION"].astype(object)

    inventory_history = read_sql(cfg.inventory_history_886)
    inventory_history.sort_values(by=["LOCATION", "ITEMNUM", "MONTH"], inplace=True)
    print(inventory_history.head())

    inventory_history["ITEMNUM"] = inventory_history["ITEMNUM"].astype(int)
    inventory_history["LOCATION"] = inventory_history["LOCATION"].astype(object)
    inventory_history = pd.pivot_table(
        inventory_history, index=["ITEMNUM", "LOCATION"], columns="MONTH", values="NET_USAGE", aggfunc=sum
    )
    inventory_history["MEAN"] = inventory_history.mean(axis=1)
    inventory_history["STDDEV"] = inventory_history.std(axis=1)
    inventory_history["CV"] = round(inventory_history["STDDEV"] / inventory_history["MEAN"], 2)
    inventory_history[inventory_history["MEAN"] < 0] = 0
    print(inventory_history.dtypes, data.dtypes)
    data = pd.merge(inventory_history, data, on=["ITEMNUM", "LOCATION"], how="left")
    data = data.drop(
        data[
            (data["EX2VMI"] == 1)
            | (data["STATUS"].eq("OBSOLETE"))
            | (data["STATUS"].eq("PENDOBS"))
            # | (data["LOCATION STATUS"].eq("OBSOLETE"))
            # | (data["LOCATION STATUS"].eq("PENDOBS"))
            | (data["CONSIGNMENT"] == 1)
            | (data["COMMODITY"].eq("CANC"))
        ].index
    )

    # take a subset of the data, we need to use the price and the quantity of each item
    data_sub = data[
        ["ITEMNUM", "DESCRIPTION", "MINLEVEL", "ORDERQTY", "CURBAL", "AVGCOST", "MEAN", "STDDEV", "CV", "STATUS"]
    ]

    # create the column of the additive cost per itemnum
    data_sub["AddCost"] = data_sub["AVGCOST"] * data_sub["MEAN"] * 12
    # order by cumulative cost
    data_sub = data_sub.sort_values(by=["AddCost"], ascending=False)
    # create the column of the running CumCost of the cumulative cost per SKU
    data_sub["RunCumCost"] = data_sub["AddCost"].cumsum()
    # create the column of the total sum
    data_sub["TotSum"] = data_sub["AddCost"].sum()
    # create the column of the running percentage
    data_sub["RunPerc"] = data_sub["RunCumCost"] / data_sub["TotSum"]
    data_sub["ABC"] = data_sub["RunPerc"].apply(abc_classification)
    data_sub["XYZ"] = data_sub["CV"].apply(xyz_classification)

    print(data_sub.head())

    # total TSNs for each class
    data_sub.ABC.value_counts()

    # total cost per class
    class_A = data_sub[data_sub.ABC == "A"]["AddCost"].count()
    class_B = data_sub[data_sub.ABC == "B"]["AddCost"].count()
    class_C = data_sub[data_sub.ABC == "C"]["AddCost"].count()
    print("Class A Items:", class_A)
    print("Class B Items:", class_B)
    print("Class C Items:", class_C)

    # total cost per class
    print("Cost of Class A :", data_sub[data_sub.ABC == "A"]["AddCost"].sum())
    print("Cost of Class B :", data_sub[data_sub.ABC == "B"]["AddCost"].sum())
    print("Cost of Class C :", data_sub[data_sub.ABC == "C"]["AddCost"].sum())

    pct_A = data_sub[data_sub.ABC == "A"]["AddCost"].sum() / data_sub["AddCost"].sum()
    pct_B = data_sub[data_sub.ABC == "B"]["AddCost"].sum() / data_sub["AddCost"].sum()
    pct_C = data_sub[data_sub.ABC == "C"]["AddCost"].sum() / data_sub["AddCost"].sum()
    # percent of total cost per class
    print("Percent of Cost of Class A :", pct_A)
    print("Percent of Cost of Class B :", pct_B)
    print("Percent of Cost of Class C :", pct_C)

    data_sub.to_csv("c:\\users\\uxkp\\desktop\\cotter.csv", index=False)
    return data_sub


def safety_stock_calculation(df, freq="weekly", service_level=0.95):
    if freq == "weekly" or freq == "w":
        avg_usage = df["WEEKLY_AVG"]
        std_usage = df["WEEKLY_STDDEV"]
        lead_time_avg = df["CALC_AVG_LT"] / 7
        lead_time_std = df["CALC_STDDEV_LT"] / 7
    else:
        avg_usage = df["WEEKLY_AVG"] / 7
        std_usage = df["WEEKLY_STDDEV"] / 7
        lead_time_avg = df["CALC_AVG_LT"]
        lead_time_std = df["CALC_STDDEV_LT"]

    lead_time_demand = avg_usage * lead_time_avg
    demand_variability = lead_time_avg * std_usage ** 2
    lead_time_variability = avg_usage ** 2 * lead_time_std ** 2
    safety_stock = norm.ppf(service_level) * np.sqrt(demand_variability + lead_time_variability)
    rop = round(safety_stock + lead_time_demand, 0)
    return rop


# TODO: economic order quantities
def rop_calculation(location):
    exceptions = pd.DataFrame(pd.read_excel("Z:\\SHAWN_MORSE\\EXCEPTIONS.xlsx", sheet_name="STANDARD QTY INFORMATION"))
    exceptions["LOCATION"] = exceptions["LOCATION"].astype(str)
    exceptions_cols = ["ITEMNUM", "LOCATION", "NOTES"]
    # load data sources into data frames (temporarily CSV files) -->

    # query_list = [cfg.inv_review_by_week, cfg.inventory_params, cfg.leadtime]
    df1 = read_sql(cfg.inv_review_by_week)  # ITEMNUM: object datatype
    # print(f"df1: {df1.dtypes}")
    # df1 = pd.read_csv("c:\\users\\uxkp\\desktop\\T&D WEEKLY NET USAGE.sql.csv")
    df1["ITEMNUM"] = df1["ITEMNUM"].astype(int)
    df1["LOCATION"] = df1["LOCATION"].astype(str)
    df2 = read_sql(cfg.inventory_params)
    # print(f"df2: {df2.dtypes}")
    # df2 = pd.read_csv("c:\\users\\uxkp\\desktop\\Inventory Parameters.sql.csv")
    df2["LOCATION"] = df2["LOCATION"].astype(str)
    df3 = read_sql(cfg.leadtime)
    # print(f"df3: {df3.dtypes}")
    # df3 = pd.read_csv("c:\\users\\uxkp\\desktop\\Lead Time Review.sql.csv")
    df4 = abc_analysis(query)
    df4 = df4[["ITEMNUM", "ABC", "XYZ"]]
    # print(f"df4: {df4.dtypes}")
    df5 = exceptions[exceptions_cols]
    # print(f"df5: {df1.dtypes}")

    #  sort historical data by location, itemnum, time -->
    df1.sort_values(by=["LOCATION", "ITEMNUM", "WEEK"], inplace=True)

    # reshape historical data -->
    df1_pivot = (
        df1.groupby(["ITEMNUM", "LOCATION"])
        .agg(
            WEEKLY_AVG=pd.NamedAgg(column="NET_USAGE", aggfunc=np.mean),
            WEEKLY_STDDEV=pd.NamedAgg(column="NET_USAGE", aggfunc=np.std),
        )
        .reset_index()
    )
    # df1_pivot["DAILY_AVG"] = df1_pivot["WEEKLY_AVG"] / 7

    # join inventory paramaters to reshaped historical data -->

    merged = pd.merge(df1_pivot, df5, on=["ITEMNUM", "LOCATION"], how="left")

    merged = merged.merge(df2, how="left", on=["ITEMNUM", "LOCATION"])
    # join lead time data to previously merged data -->

    merged = merged.merge(df3, how="left", on="ITEMNUM")

    merged = merged.merge(df4, how="left", on="ITEMNUM")

    stdpk_regex = r"((?<!\S)(^\/)?(\b[0-9]+\s+\w{2}|[0-9]+)\/+([a-zA-Z]{2})\b)"
    minlevel = merged["MINLEVEL"]

    regex_df = merged["DESCRIPTION"].str.extractall(stdpk_regex)
    regex_df = regex_df.groupby(level=0)[0].apply(list)
    merged["STD_PKG"] = regex_df

    merged.loc[merged["CALC_AVG_LT"].isnull(), "CALC_AVG_LT"] = merged["MAXIMO_LT"]
    merged.loc[merged["CALC_STDDEV_LT"].isnull(), "CALC_STDDEV_LT"] = round(merged["MAXIMO_LT"] * 0.2, 0)

    high_certainty = safety_stock_calculation(merged, "d", service_level=0.95)
    med_certainty = safety_stock_calculation(merged, "d", service_level=0.85)
    low_certainty = safety_stock_calculation(merged, "d", service_level=0.65)
    merged.loc[merged["ABC"].eq("A") & merged["XYZ"].eq("X"), "CALCULATED ROP"] = high_certainty
    merged.loc[
        (merged["ABC"].ne("A") & merged["XYZ"].eq("X")) | (merged["ABC"].eq("A") & merged["XYZ"].ne("X")),
        "CALCULATED ROP",
    ] = med_certainty
    merged.loc[merged["ABC"].ne("A") & merged["XYZ"].ne("X"), "CALCULATED ROP"] = low_certainty

    merged["CALCULATED COST DELTA"] = (merged["CALCULATED ROP"] - merged["MINLEVEL"]) * merged["AVGCOST"]

    # ***********************************************************************************************
    ## higher precedent -->
    c00 = "exceptions list"
    c01 = "Do not use"
    c02 = "excluded - CANC"
    c03 = "Obsolete - should be CANC"
    c04 = "SET ROP/EOQ AT -1/1 (PENDOBS)"

    # comments and exceptions -->

    a00 = "excluded - REORDER flag OFF"

    ## lower importance -->
    e00 = "excluded - Consignment"
    e01 = "excluded - Steel Reel"
    e02 = "excluded - VMI Item"
    e03 = "ON HAND & NO USAGE 12-MONTHS"
    e04 = "excluded - Tools"
    e05 = "excluded - Safety Item"
    e06 = "last transaction date before 2017"
    e07 = r"excluded - TELECOM \ ROUTERS"
    e09 = "excluded - not in ACTIVE status"

    merged.loc[merged["REORDER"].eq(0), "COMMENT"] = a00
    merged.loc[merged["CONSIGNMENT"].eq(1), "COMMENT"] = e00
    merged.loc[merged["COMMODITY"].eq("2830"), "COMMENT"] = e01
    merged.loc[merged["EX2VMI"].eq(1), "COMMENT"] = e02
    merged.loc[merged["COMMODITY"].isin(["0804", "0806"]), "COMMENT"] = e04
    merged.loc[merged["COMMODITYGROUP"].eq("23"), "COMMENT"] = e05
    merged.loc[merged["COMMODITY"].eq("1812"), "COMMENT"] = e07
    merged.loc[merged["STATUS"].ne("ACTIVE"), "COMMENT"] = e09

    merged.loc[merged["NOTES"].notnull(), "COMMENT"] = c00
    merged.loc[merged["DESCRIPTION"].str.contains("do not use", case=False, na=False), "COMMENT"] = c01
    merged.loc[merged["MINLEVEL"].ne(-1) & merged["STATUS"].eq("PENDOBS"), "COMMENT"] = c02
    merged.loc[merged["DESCRIPTION"].str.contains("obsolete", case=False, na=False), "COMMENT"] = c03
    merged.loc[merged["COMMODITY"].eq("CANC"), "COMMENT"] = c04

    # after necessary adjustments to calculated rop, instantiate recommended rop
    # handle exceptions (such as VMI, Consignment, etc)
    merged["RECOMMENDED ROP"] = merged["CALCULATED ROP"]
    merged.loc[merged["COMMENT"].isin([c01, c02, c03, c04, e00, e01, e02, e09]), "RECOMMENDED ROP"] = -1

    merged.loc[
        minlevel.eq(-1)
        | merged["AVLBLBALANCE"].lt(0)
        | merged["AVGCOST"].gt(750)
        | merged["COMMENT"].isin([c00, e03, e04, e05, e06, e07]),
        "RECOMMENDED ROP",
    ] = minlevel
    merged["RECOMMENDED COST DELTA"] = (merged["RECOMMENDED ROP"] - merged["MINLEVEL"]) * merged["AVGCOST"]
    # ***********************************************************************************************

    # TODO: create function to simplify input/output data

    # output finished dataframe to csv file
    finished_report = merged[
        [
            "ITEMNUM",
            "DESCRIPTION",
            "LOCATION",
            "MINLEVEL",
            "ORDERQTY",
            "RECOMMENDED ROP",
            "RECOMMENDED COST DELTA",
            "WEEKLY_AVG",
            "WEEKLY_STDDEV",
            "MAXIMO_LT",
            "CALC_AVG_LT",
            "CALC_STDDEV_LT",
            "CURBAL",
            "RESERVEDQTY",
            "AVLBLBALANCE",
            "AVGCOST",
            "EXTENDED_COST",
            "CALCULATED ROP",
            "CALCULATED COST DELTA",
            "COMMODITY",
            "COMMODITYGROUP",
            "COMMODITYDESC",
            "RECEIPT_TXNS",
            "LASTISSUEDATE",
            "LASTTRANSFERDATE",
            "CONSIGNMENT",
            "EX2VMI",
            "REORDER",
            "LOWCOST",
            "INTERNAL",
            "STATUS",
            "LOCATION STATUS",
            "STORELOC",
            "STD_PKG",
            "COMMENT",
            "NOTES",
            "ABC",
            "XYZ",
        ]
    ]

    finished_report.to_csv(f"c:\\users\\uxkp\\desktop\\{location}_ROP_review_{timestring}.csv", index=False)


if __name__ == "__main__":
    rop_calculation("886")
