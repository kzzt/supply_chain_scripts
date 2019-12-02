import cx_Oracle
import pandas as pd
from sqlalchemy import create_engine

import onccfg


def oracle_conn():
    user = onccfg.username
    password = onccfg.password
    sid = cx_Oracle.makedsn(onccfg.host, onccfg.port, sid=onccfg.sid)  # host, port, sid
    cstr = f"oracle://{user}:{password}@{sid}"
    engine = create_engine(cstr, convert_unicode=False, pool_recycle=10, pool_size=50, echo=False)
    return engine


def read_sql(filename):
    engine = oracle_conn()
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()
    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))
    df_from_sql.columns = map(str.upper, df_from_sql.columns)
    return df_from_sql


def data_prep(source_data):
    df = source_data
    df.rename(columns={df.columns[0]: "DATETIME", df.columns[1]: "ITEMNUM"}, inplace=True)
    df = df.set_index(["DATETIME", "ITEMNUM"])
    df.update(df.groupby(level="ITEMNUM").bfill())
    return df


def daily_time_series(clean_dataframe):
    df = (
        clean_dataframe.reset_index()
            .pivot_table(index="DATETIME", columns="ITEMNUM", values="INVENTORY", aggfunc=sum)
            .bfill()
            .unstack()
            .reset_index(name="INVENTORY")
    )

    print(df.head(), df.tail())

    sum_df = df.groupby(df["DATETIME"]).sum().reset_index()
    sum_df = sum_df[["DATETIME", "INVENTORY"]]

    print(sum_df.head(), sum_df.tail())

    sum_df.to_csv(r"c:\users\uxkp\desktop\time_series_test_886.csv")


sql = r"C:\Users\uxkp\sql_queries\inventory\avg_bal.sql"
daily_time_series(data_prep(read_sql(sql)))
