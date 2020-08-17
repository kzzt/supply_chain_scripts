import time

import cx_Oracle
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sqlalchemy import create_engine

import onccfg as cfg

user = cfg.username
password = cfg.password
sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

cstr = f"oracle://{user}:{password}@{sid}"

engine = create_engine(
    cstr, convert_unicode=False, pool_recycle=10, pool_size=50, echo=False
)


def read_sql(filename):
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()

    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)

    return df_from_sql


def abc_classification(perc):
    if perc > 0 and perc < 0.8:
        return "A"
    elif perc >= 0.8 and perc < 0.9:
        return "B"
    elif perc >= 0.9:
        return "C"


def xyz_classification(cv):
    if cv > 0 and cv < 0.5:
        return "X"
    elif cv >= 0.5 and cv < 1.5:
        return "Y"
    elif cv >= 1.5:
        return "Z"


query = read_sql("D:\\SQL_Queries\\ABC_CLASS_PARAMETERS.sql")


def abc_analysis(data: object, report_type: object = "Test") -> object:
    data["ITEMNUM"] = data["ITEMNUM"].astype(int)
    data["LOCATION"] = data["LOCATION"].astype(object)

    inventory_history = read_sql(cfg.trinity_hist)
    inventory_history.sort_values(by=["LOCATION", "ITEMNUM", "MONTH"], inplace=True)
    print(inventory_history.head())
    inventory_history = pd.pivot_table(
        inventory_history,
        index=["ITEMNUM", "LOCATION"],
        columns="MONTH",
        values="NET_USAGE",
        aggfunc=sum,
    )

    inventory_history["MEAN"] = inventory_history.mean(axis=1)
    inventory_history["STDDEV"] = inventory_history.std(axis=1)
    inventory_history["CV"] = round(
        inventory_history["STDDEV"] / inventory_history["MEAN"], 2
    )
    inventory_history[inventory_history["MEAN"] < 0] = 0
    print(
        f"data types: {data.dtypes} --- inv_history_types: {inventory_history.dtypes}"
    )
    data = pd.merge(inventory_history, data, on=["ITEMNUM", "LOCATION"], how="left")

    # EXCLUDE: obsolete, VMI, consignment from classification
    # vmi = data[data["EX2VMI"] == 1].index
    # data.drop(vmi, inplace=True)
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
        [
            "ITEMNUM",
            "DESCRIPTION",
            "MINLEVEL",
            "ORDERQTY",
            "CURBAL",
            "AVGCOST",
            "MEAN",
            "STDDEV",
            "CV",
            "STATUS",
        ]
    ]

    # create the column of the additive cost per itemnum
    data_sub["AddCost"] = data_sub["AVGCOST"] * data_sub["MEAN"] * 12
    data_sub.loc[data_sub['AddCost']] = data_sub["AVGCOST"] * data_sub["MEAN"] * 12
    # order by cumulative cost
    data_sub = data_sub.sort_values(by=["AddCost"], ascending=False)
    # create the column of the running CumCost of the cumulative cost per SKU
    data_sub["RunCumCost"] = data_sub["AddCost"].cumsum()
    # create the column of the total sum
    data_sub["TotSum"] = data_sub["AddCost"].sum()
    # create the column of the running percentage
    data_sub["RunPerc"] = data_sub["RunCumCost"] / data_sub["TotSum"]
    # create the column of the class
    # data_sub.loc[data_sub["EX2VMI"].eq(0), "Class"] = data_sub["RunPerc"].apply(abc_classification)
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

    performance = data_sub["AddCost"].tolist()
    y_pos = np.arange(len(performance))

    plt.plot(y_pos, performance)
    plt.ylabel("Cost")
    plt.title("ABC Analysis - Cost per TSN")
    plt.grid(True)
    plt.ylim((0, 100_000))
    plt.show()

    fig = plt.figure()
    ax = fig.add_subplot(111)
    a = ax.annotate(
        "A",
        xy=(class_A, pct_A),
        xytext=(class_A + 1000, 0.75),
        arrowprops=dict(facecolor="green", shrink=0.05),
    )
    b = ax.annotate(
        "B",
        xy=(class_B + class_B, pct_A + pct_B),
        xytext=(class_B + class_A + 1000, 0.85),
        arrowprops=dict(facecolor="blue", shrink=0.05),
    )

    c = ax.annotate(
        "C",
        xy=(class_C + class_B + class_A, pct_A + pct_B + pct_C),
        xytext=(class_C + class_B + class_A - 1000, 0.91),
        arrowprops=dict(facecolor="red", shrink=0.05),
    )
    scalar_A = pct_A
    scalar_B = pct_B + pct_A
    scalar_C = pct_C + pct_B + pct_A
    plt.vlines(class_A, [0], scalar_A, colors="green", linestyles="dotted")
    plt.vlines(class_B + class_A, [0], scalar_B, colors="blue", linestyles="dotted")
    plt.vlines(
        class_C + class_B + class_A, [0], scalar_C, colors="red", linestyles="dotted"
    )
    performance = data_sub["RunPerc"].tolist()
    y_pos = np.arange(len(performance))

    plt.plot(y_pos, performance)
    plt.ylabel("Running Total Percentage")
    plt.title("ABC Analysis - Cumulative Cost per ITEMNUM")
    plt.grid(True)
    plt.show()

    timestring = time.strftime("%Y-%m-%d")
    data_sub.to_csv("c:\\users\\u6zn\\desktop\\abc_analysis.csv", index=False)
    writer = pd.ExcelWriter(
        f"C:\\Users\\u6zn\\Desktop\\ABC CLASSIFICATION {timestring}.xlsx"
    )

    data_sub.to_excel(writer, "ABC CLASSIFICATION", index=False)
    writer.save()
    return data_sub
