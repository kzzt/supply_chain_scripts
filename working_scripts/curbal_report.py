# queries VMI po from the previous week and outputs to file named for PONUMs
import os
import random
from datetime import datetime, date, time, timedelta
from time import mktime as mk

import cx_Oracle
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

# sql_file = open(cfg.vmi_query, "r")
# vmi_query = sql_file.read()
# sql_file.close()

query_string = (
    "SELECT ITEMNUM, LOCATION, SUM(CURBAL) AS CURBAL FROM MSCRADS.INVBALANCES group by ITEMNUM, LOCATION"
)



def timeFix(filename):
    hours = time(7, random.randrange(0, 29))
    today = date.today()
    last_monday = today - timedelta(days=today.weekday())

    timestamp = mk(datetime.combine(last_monday, hours).timetuple())
    os.utime(
        f"Z:\\SHAWN_MORSE\\VMI PO Validation\\{filename}.xlsx",
        (timestamp, timestamp + 10),
    )


def main():
    curbal_df = pd.DataFrame(pd.read_sql_query(query_string, engine))
    curbal_df.columns = map(str.upper, curbal_df.columns)

    curbal_df = curbal_df.pivot(index="ITEMNUM", columns="LOCATION", values="CURBAL").reset_index()
    curbal_df['ITEMNUM'] = curbal_df['ITEMNUM'].astype(int)

    backorder_df = pd.read_csv("c:\\users\\u6zn\\downloads\\VelocityBackorder (1).csv")
    backorder_html = pd.DataFrame(pd.read_html(r"C:\Users\u6zn\PycharmProjects\supply_chain_scripts\working_scripts\velocity-table.html"))
    print(backorder_html.head())


    new_df = pd.merge(backorder_df, curbal_df, how="left", on="ITEMNUM")

    # filename_str = " & ".join(map(str, curbal_df["PONUM"].unique()))

    # print(filename_str)

    writer = pd.ExcelWriter(f"c:\\users\\u6zn\\desktop\\curbal_6.xlsx")

    new_df.to_excel(writer, "ON HAND", index=False)
    workbook = writer.book
    worksheet = writer.sheets["ON HAND"]
    # worksheet.set_column("F:F", 32)

    writer.save()

    # timeFix(filename_str)


if __name__ == "__main__":
    main()
