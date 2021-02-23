import time
import cx_Oracle
import numpy as np
import pandas as pd
from scipy.stats import norm
from sqlalchemy import create_engine
import abc_classification
import onccfg as cfg

timestring = time.strftime("%Y-%m-%d")


def oracle_conn():
    user = cfg.username
    password = cfg.password
    sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

    cstr = f"oracle://{user}:{password}@{sid}"

    engine = create_engine(
        cstr, convert_unicode=False, pool_recycle=20, pool_size=100, echo=False
    )
    return engine


def read_sql(filename):
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()

    engine = oracle_conn()

    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)

    return df_from_sql


# def abc_classification(perc):
#     if 0 < perc < 0.8:
#         return "A"
#     elif 0.8 <= perc < 0.9:
#         return "B"
#     elif perc >= 0.9:
#         return "C"
#
#
# def xyz_classification(cv):
#     if 0 < cv < 0.5:
#         return "X"
#     elif 0.5 <= cv < 1.5:
#         return "Y"
#     elif cv >= 1.5:
#         return "Z"


query = read_sql("D:\\SQL_Queries\\ABC_CLASS_PARAMETERS.sql")


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
    safety_stock = norm.ppf(service_level) * np.sqrt(
        demand_variability + lead_time_variability
    )
    rop = round(safety_stock + lead_time_demand, 0)
    return rop


# TODO: economic order quantities
def rop_calculation(location):
    # load data sources into data frames (temporarily CSV files) -->
    df1 = read_sql(cfg.trinity_wkly)  # ITEMNUM: object datatype
    df1["ITEMNUM"] = df1["ITEMNUM"].astype(int)
    df1["LOCATION"] = df1["LOCATION"].astype(str)
    df2 = read_sql(cfg.inventory_params)
    df2["LOCATION"] = df2["LOCATION"].astype(str)
    df3 = read_sql(cfg.leadtime)
    df4 = abc_classification.abc_analysis(abc_classification.query, "TEST")
    df4 = df4[["ITEMNUM", "ABC", "XYZ"]]

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

    # join inventory paramaters to reshaped historical data -->

    # merged = pd.merge(df1_pivot, df5, on=["ITEMNUM", "LOCATION"], how="left")

    merged = df1.merge(df2, how="left", on=["ITEMNUM", "LOCATION"])
    # join lead time data to previously merged data -->

    merged = merged.merge(df3, how="left", on="ITEMNUM")

    merged = merged.merge(df4, how="left", on="ITEMNUM")

    stdpk_regex = r"((?<!\S)(^\/)?(\b[0-9]+\s+\w{2}|[0-9]+)\/+([a-zA-Z]{2})\b)"
    minlevel = merged["MINLEVEL"]

    regex_df = merged["DESCRIPTION"].str.extractall(stdpk_regex)
    regex_df = regex_df.groupby(level=0)[0].apply(list)
    merged["STD_PKG"] = regex_df

    merged.loc[merged["CALC_AVG_LT"].isnull(), "CALC_AVG_LT"] = merged["MAXIMO_LT"]
    merged.loc[merged["CALC_STDDEV_LT"].isnull(), "CALC_STDDEV_LT"] = round(
        merged["MAXIMO_LT"] * 0.2, 0
    )

    high_certainty = safety_stock_calculation(merged, "d", service_level=0.95)
    med_certainty = safety_stock_calculation(merged, "d", service_level=0.85)
    low_certainty = safety_stock_calculation(merged, "d", service_level=0.65)
    merged.loc[
        merged["ABC"].eq("A") & merged["XYZ"].eq("X"), "CALCULATED ROP"
    ] = high_certainty
    merged.loc[
        (merged["ABC"].ne("A") & merged["XYZ"].eq("X"))
        | (merged["ABC"].eq("A") & merged["XYZ"].ne("X")),
        "CALCULATED ROP",
    ] = med_certainty
    merged.loc[
        merged["ABC"].ne("A") & merged["XYZ"].ne("X"), "CALCULATED ROP"
    ] = low_certainty

    merged["CALCULATED COST DELTA"] = (
        merged["CALCULATED ROP"] - merged["MINLEVEL"]
    ) * merged["AVGCOST"]

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
    merged.loc[
        merged["DESCRIPTION"].str.contains("do not use", case=False, na=False),
        "COMMENT",
    ] = c01
    merged.loc[
        merged["MINLEVEL"].ne(-1) & merged["STATUS"].eq("PENDOBS"), "COMMENT"
    ] = c02
    merged.loc[
        merged["DESCRIPTION"].str.contains("obsolete", case=False, na=False), "COMMENT"
    ] = c03
    merged.loc[merged["COMMODITY"].eq("CANC"), "COMMENT"] = c04

    # after necessary adjustments to calculated rop, instantiate recommended rop
    # handle exceptions (such as VMI, Consignment, etc)
    merged["RECOMMENDED ROP"] = merged["CALCULATED ROP"]
    merged.loc[
        merged["COMMENT"].isin([c01, c02, c03, c04, e00, e01, e02, e09]),
        "RECOMMENDED ROP",
    ] = -1

    merged.loc[
        minlevel.eq(-1)
        | merged["AVLBLBALANCE"].lt(0)
        | merged["AVGCOST"].gt(750)
        | merged["COMMENT"].isin([c00, e03, e04, e05, e06, e07]),
        "RECOMMENDED ROP",
    ] = minlevel
    # --- rounding -->
    merged.loc[merged["RECOMMENDED ROP"].gt(5000), "RECOMMENDED ROP"] = round(
        merged["RECOMMENDED ROP"], -3
    )
    merged.loc[merged["RECOMMENDED ROP"].gt(500), "RECOMMENDED ROP"] = round(
        merged["RECOMMENDED ROP"], -2
    )
    merged.loc[merged["RECOMMENDED ROP"].gt(50), "RECOMMENDED ROP"] = round(
        merged["RECOMMENDED ROP"], -1
    )
    # <-- rounding ---
    merged["RECOMMENDED COST DELTA"] = (
        merged["RECOMMENDED ROP"] - merged["MINLEVEL"]
    ) * merged["AVGCOST"]
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

        f"c:\\users\\u6zn\\desktop\\{location}_ROP_review_{timestring}.csv", index=False
    )


if __name__ == "__main__":
    rop_calculation("886")