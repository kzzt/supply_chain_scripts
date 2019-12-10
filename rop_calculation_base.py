import time

import cx_Oracle
import numpy as np
import pandas as pd
from scipy.stats import norm
from sqlalchemy import create_engine

import onccfg as cfg

start = time.process_time()

user = cfg.username
password = cfg.password
sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

cstr = f"oracle://{user}:{password}@{sid}"

engine = create_engine(cstr, convert_unicode=False, pool_recycle=10, pool_size=50, echo=False)


# excel file in which to save report


# function to read sql query file, run query, and produce data-frame from queried data
def read_sql(filename):
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()

    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)

    return df_from_sql


def apics_calculation(df, divisor=7, service_level=0.90):
    """Calculate safety stock and reorder point using APICS method.
    Parameters
    :param df: source dataframe
    :param divisor: 1 for weekly, 7 for daily, 30 for monthly
    :param z_value: default 1.65
    ----------
    returns
    Integer with reorder point from data set.
    """

    lead_time_avg = df["AVG_LT"] / divisor
    lead_time_stddev = df["STDDEV_LT"] / divisor

    if divisor is 1:
        demand_variation = lead_time_avg * df["DAILY_STDDEV"] ** 2
        lead_time_variation = (lead_time_stddev * df["DAILY_AVG"]) ** 2
    else:
        demand_variation = lead_time_avg * df["WEEKLY_STDDEV"] ** 2
        lead_time_variation = (lead_time_stddev * df["WKLY_AVG"]) ** 2

    leadtime_demand = df["DAILY_AVG"] * lead_time_avg
    reorder_point = round(
        np.sqrt(demand_variation + lead_time_variation) * norm.ppf(service_level) + leadtime_demand, 0
    )

    return reorder_point


def vermorel_calculation(df, service_level=0.95):  # basic rop calc using only demand uncertainty
    lead_time_avg = df["AVG_LT"]
    leadtime_demand = df["DAILY_AVG"] * lead_time_avg
    rop = round((df["DAILY_STDDEV"] * norm.ppf(service_level) * np.sqrt(df["AVG_LT"])) + leadtime_demand, 0)
    return rop


def main(query, type="Test"):
    # for filename date
    timestring = time.strftime("%Y-%m-%d")
    writer = pd.ExcelWriter(f"C:\\Users\\UXKP\\Desktop\\886_ROP_{type}_" + timestring + ".xlsx")

    # run query for relevant locations
    inventory_history = read_sql(query)

    exceptions = pd.DataFrame(pd.read_excel(r"Z:\SHAWN_MORSE\EXCEPTIONS.xlsx"))
    exceptions["LOCATION"] = exceptions["LOCATION"].astype(str)
    exceptions = exceptions.loc[exceptions["LOCATION"].eq("886"), ["ITEMNUM", "LOCATION", "NOTES"]]

    # exceptions list combined with inventory history

    inventory_history = pd.merge(inventory_history, exceptions, on=["ITEMNUM", "LOCATION"], how="left")

    inventory_history["APICS ROP"] = apics_calculation(inventory_history, divisor=1, service_level=0.90)
    # inventory_history["VERMOREL ROP"] = vermorel_calculation(inventory_history, service_level=0.96)

    # if calculated to 0 and existing ROP is -1 then keep at -1
    inventory_history.loc[inventory_history["APICS ROP"].eq(0) & inventory_history["MINLEVEL"].eq(-1), "APICS ROP"] = -1
    inventory_history.loc[
        inventory_history["VERMOREL ROP"].eq(0) & inventory_history["MINLEVEL"].eq(-1), "VERMOREL ROP"
    ] = -1

    # stdpk_regex = r"(\d*\s\w*\/[a-zA-Z]{2})"
    # stdpk_regex = r"\s(^\/)?(\b[0-9]+\s+\w{2}|[0-9]+)\/+([a-zA-Z]{2})\b"
    stdpk_regex = r"((?<!\S)(^\/)?(\b[0-9]+\s+\w{2}|[0-9]+)\/+([a-zA-Z]{2})\b)"
    minlevel = inventory_history["MINLEVEL"]

    regex_df = inventory_history["DESCRIPTION"].str.extractall(stdpk_regex)
    regex_df = regex_df.groupby(level=0)[0].apply(list)
    inventory_history["STD_PKG"] = regex_df

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

    # date_cutoff_limit = datetime.date.today() - datetime.date(timedelta(days=730))

    inventory_history.loc[inventory_history["REORDER"].eq(0), "COMMENT"] = a00
    inventory_history.loc[inventory_history["CONSIGNMENT"].eq(1), "COMMENT"] = e00
    inventory_history.loc[inventory_history["COMMODITY"].eq("2830"), "COMMENT"] = e01
    inventory_history.loc[inventory_history["EX2VMI"].eq(1), "COMMENT"] = e02
    # inventory_history.loc[(lid.dt.year < 2017 | lid.isnull()) & (ltd.dt.year < 2017 | ltd.isnull()), "COMMENT"] = e06
    inventory_history.loc[inventory_history["COMMODITY"].isin(["0804", "0806"]), "COMMENT"] = e04
    inventory_history.loc[inventory_history["COMMODITYGROUP"].eq("23"), "COMMENT"] = e05
    inventory_history.loc[inventory_history["COMMODITY"].eq("1812"), "COMMENT"] = e07
    inventory_history.loc[inventory_history["STATUS"].ne("ACTIVE"), "COMMENT"] = e09

    inventory_history.loc[inventory_history["NOTES"].notnull(), "COMMENT"] = c00
    inventory_history.loc[
        inventory_history["DESCRIPTION"].str.contains("do not use", case=False, na=False), "COMMENT"
    ] = c01
    inventory_history.loc[
        inventory_history["MINLEVEL"].ne(-1) & inventory_history["STATUS"].eq("PENDOBS"), "COMMENT"
    ] = c02
    inventory_history.loc[
        inventory_history["DESCRIPTION"].str.contains("obsolete", case=False, na=False), "COMMENT"
    ] = c03
    inventory_history.loc[inventory_history["COMMODITY"].eq("CANC"), "COMMENT"] = c04

    # after necessary adjustments to calculated rop, instantiate recommended rop
    # handle exceptions (such as VMI, Consignment, etc)
    inventory_history["RECOMMENDED ROP"] = inventory_history["VERMOREL ROP"]
    inventory_history.loc[
        inventory_history["COMMENT"].isin([c01, c02, c03, c04, e00, e01, e02, e09]), "RECOMMENDED ROP"
    ] = -1

    inventory_history.loc[
        minlevel.eq(-1)
        | inventory_history["AVLBLBALANCE"].lt(0)
        | inventory_history["AVGCOST"].gt(750)
        | inventory_history["COMMENT"].isin([c00, e03, e04, e05, e06, e07]),
        "RECOMMENDED ROP",
    ] = minlevel
    # calculate eoq

    # calculate changes to maxlevel cost

    current_max_value = (inventory_history["MINLEVEL"] + inventory_history["ORDERQTY"]) * inventory_history["AVGCOST"]
    rec_max_value = (inventory_history["RECOMMENDED ROP"] + inventory_history["ORDERQTY"]) * inventory_history[
        "AVGCOST"
    ]
    calc_max_value = (inventory_history["VERMOREL ROP"] + inventory_history["ORDERQTY"]) * inventory_history["AVGCOST"]
    inventory_history["RECOMMENDED COST DELTA"] = rec_max_value - current_max_value
    inventory_history["CALCULATED COST DELTA"] = calc_max_value - current_max_value
    inventory_history["CURRENT DAYS ON HAND"] = round(inventory_history["CURBAL"] / inventory_history["DAILY_AVG"], 0)
    inventory_history["VERMOREL ROP DAYS ON HAND"] = round(
        (inventory_history["VERMOREL ROP"] + inventory_history["ORDERQTY"]) / inventory_history["DAILY_AVG"], 0
    )

    # df_columns = inventory_history[
    #     "SITEID",
    #     "LOCATION",
    #     "ITEMNUM",
    #     "DESCRIPTION",
    #     "COMMODITYDESC",
    #     "MINLEVEL",
    #     "ORDERQTY",
    #     "APICS ROP",
    #     "RECOMMENDED ROP",
    #     "CURBAL",
    #     "AVGCOST",
    #     "ONHAND_VALUE",
    #     "ISSUES",
    #     "TRANSFERS_OUT",
    #     "NET_USAGE",
    #     "WKLY_AVG",
    #     "RECOMMENDED COST DELTA",
    #     "CALCULATED COST DELTA",
    #     "WEEKLY_STDDEV",
    #     "RECEIPT_TXNS",
    #     "DAILY_AVG",
    #     "DAILY_STDDEV",
    #     "AVG_LT",
    #     "STDDEV_LT",
    #     "MAXIMO_LT",
    #     "EX2VMI",
    #     "REORDER",
    #     "STATUS",
    #     "COMMENT", "DAYS ON HAND"
    # ]

    # write to excel file (separate sheets for different locations)
    inventory_history.to_excel(writer, "ROP CALCULATION", index=False)

    writer.save()

    print(time.process_time() - start)


if __name__ == "__main__":
    main(cfg.inv_review_by_week)
