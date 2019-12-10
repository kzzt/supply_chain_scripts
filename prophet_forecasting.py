import time
import warnings

import cx_Oracle
import numpy as np
import pandas as pd
from fbprophet import Prophet
from fbprophet.diagnostics import cross_validation, performance_metrics
from fbprophet.plot import add_changepoints_to_plot
from fbprophet.plot import plot_cross_validation_metric
from matplotlib import pyplot as plt
from sqlalchemy import create_engine

import onccfg

start = time.process_time()
timestring = time.strftime("%Y-%m-%d")


def oracle_conn():
    starttime = time.process_time()
    warnings.filterwarnings("ignore", "Could not determine compatibility version")
    user = onccfg.username
    password = onccfg.password
    sid = cx_Oracle.makedsn(onccfg.host, onccfg.port, sid=onccfg.sid)  # host, port, sid
    cstr = f"oracle://{user}:{password}@{sid}"
    engine = create_engine(cstr, convert_unicode=False, pool_recycle=10, pool_size=50, echo=False)
    print("oracle connection time: ", time.process_time() - starttime)
    return engine


def read_sql_to_df(filename):
    starttime = time.process_time()
    engine = oracle_conn()
    sql_file = open(filename, "r")
    query = sql_file.read()
    sql_file.close()
    df_from_sql = pd.DataFrame(pd.read_sql_query(query, engine))

    df_from_sql.to_csv(f"c:\\users\\uxkp\desktop\\raw_data_PLACEHOLDER_issues_{timestring}.csv", index=False)
    print("query to dataframe time: ", time.process_time() - starttime)

    return df_from_sql


def data_prep(source_data):
    df = source_data
    df.sort_values(by=["location", "itemnum", "month"], inplace=True)
    df.columns = map(str.upper, df.columns)
    df.rename(columns={df.columns[0]: "DATETIME", df.columns[1]: "ITEMNUM"}, inplace=True)
    df = df.set_index(["DATETIME", "ITEMNUM"])
    df.update(df.groupby(level="ITEMNUM").bfill())
    print(df.columns)
    return df


def daily_time_series(clean_dataframe):
    clean_dataframe.columns = map(str.upper, clean_dataframe.columns)
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
    return sum_df


def main(timey, fcst_type):
    location = "886-limited"
    timestring = time.strftime("%Y-%m-%d")
    sql = r"C:\Users\uxkp\sql_queries\inventory\avg_bal_daily.sql"
    df = daily_time_series(data_prep(read_sql_to_df(sql)))
    print(df.columns)
    df.rename(columns={"DATETIME": "ds", "INVENTORY": "y"}, inplace=True)
    print(df.dtypes)
    print(df.tail)

    df.reset_index(drop=True)

    # m = Prophet()

    df["y"] = np.log(df["y"].replace(0, np.nan))
    df.to_csv(r"c:\users\uxkp\desktop\filetest.csv", index=False)

    df["cap"] = 24.5
    df["floor"] = 0

    m = Prophet(
        mcmc_samples=100,
        growth="logistic",
        uncertainty_samples=100,
        interval_width=0.9,
        changepoint_range=0.90,
        changepoint_prior_scale=0.01,
        weekly_seasonality=False,
    )
    m.fit(df)

    future = m.make_future_dataframe(periods=60, freq="d")
    future["cap"] = 24.5
    future["floor"] = 0
    forecast = m.predict(future)
    psf = m.predictive_samples(future)
    # m.plot(psf).show()

    forecast.to_csv(f"c:\\users\\uxkp\\desktop\\{timey}_{fcst_type}_forecast_{location}_{timestring}.csv", index=False)

    dcv = cross_validation(m, initial="30 days", period="7 days", horizon="60 days")

    performance_metrics_results = performance_metrics(dcv)
    performance_metrics_results.to_csv(
        f"c:\\users\\uxkp\\desktop\\{timey}_{fcst_type}_performance_metrics_{location}_{timestring}.csv", index=False
    )

    fig1 = plot_cross_validation_metric(dcv, metric="mape")
    plt.savefig(
        f"c:\\users\\uxkp\\desktop\\{timey}_{fcst_type}_cross validation_{location}_{timestring}.png", format="png"
    )
    fig2 = m.plot(forecast, xlabel=f"Time Series Forecast")
    plt.savefig(f"c:\\users\\uxkp\\desktop\\{timey}_{fcst_type}_forecast_{location}_{timestring}.png", format="png")
    fig3 = m.plot_components(forecast)
    a = add_changepoints_to_plot(fig3.gca(), m, forecast)
    plt.savefig(f"c:\\users\\uxkp\\desktop\\{timey}_{fcst_type}_components_{location}_{timestring}.png", format="png")
    fig1.show()
    fig2.show()
    fig3.show()


if __name__ == "__main__":
    main("daily", "TEST")
    print("TOTAL time: ", time.process_time() - start)
    # for s in site:
    #     main("daily", s, "issues")
    #     print("TOTAL time: ", time.process_time() - start)
