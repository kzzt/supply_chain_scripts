# queries VMI po from the previous week and outputs to file named for PONUMs
from xlrd import open_workbook

import connect as ct

import time
import pandas as pd
import numpy as np

import win32com.client as win32

filename = "c:\\users\\u6zn\\downloads\\WO_Shipment_Update_20210717_230002.dat"

df = pd.read_csv(filename, skiprows=1)
df = df.loc[df['SHIP_STATUS'] == 'LANE']


work_orders = ct.irby_workorders()

cost_query = ("SELECT\n"
              "	WO.PARENT,\n"
              "	WP.WONUM, WP.LOCATION , SHIPTO.SHIPTO, SHIPTO.SHIPTONAME,\n"
              "	WP.ITEMNUM,\n"
              "	WP.ITEMQTY ,\n"
              "	WP.LINECOST ,\n"
              "	WP.UNITCOST, U1.MEMO\n"
              "FROM\n"
              "	MXRADS.WPMATERIAL WP\n"
              "LEFT JOIN (\n"
              "	SELECT\n"
              "		WONUM,\n"
              "		PARENT,\n"
              "		ESTMATCOST\n"
              "	FROM\n"
              "		MXRADS.WORKORDER\n"
              "	WHERE\n"
              f"		WONUM IN ({work_orders})) WO ON\n"
              "	WO.WONUM = WP.WONUM\n"
              "	LEFT JOIN (SELECT ITEMNUM, STORELOC, QUANTITY, PONUM, REFWO, LOCATION, MEMO FROM MSCRADS.MATUSETRANS) U1 ON U1.ITEMNUM = WP.ITEMNUM AND U1.STORELOC = WP.LOCATION AND U1.REFWO = WP.WONUM\n"
              f"LEFT JOIN (SELECT WONUM, DESCRIPTION AS SHIPTONAME, SADDRESSCODE AS SHIPTO FROM MXRADS.WOSERVICEADDRESS WHERE WONUM IN ({work_orders})) SHIPTO ON SHIPTO.WONUM = WP.WONUM\n"
              "	\n"
              "	WHERE\n"
              f"	WP.WONUM IN ({work_orders})\n"
              "ORDER BY\n"
              "	WO.PARENT,\n"
              "	WP.WONUM,\n"
              "	WP.ITEMNUM ")

# cost_data_sql = "C:\\Users\\u6zn\\AppData\\Roaming\\DBeaverData\\workspace6\\General\\Scripts\\IRBY_COST.sql"


timestring = time.strftime("%m.%d.%y")


def main():
    cost_df = pd.DataFrame(pd.read_sql_query(cost_query, ct.connect()))
    cost_df.columns = map(str.upper, cost_df.columns)
    print(cost_df.dtypes)
    cost_df["WONUM"] = cost_df["WONUM"].astype(int)

    irby_workbook = pd.ExcelWriter(
        f"c:\\users\\u6zn\\desktop\\Irby WOs for Audit {timestring}.xlsx",
        engine="xlsxwriter",
    )

    cost_df.to_excel(irby_workbook, sheet_name="Irby WOs", index=False)

    cost_df_pivot = (
        cost_df.groupby(["PARENT", "WONUM", "LOCATION", "SHIPTO"])
        .agg(ESTMATCOST=pd.NamedAgg(column="LINECOST", aggfunc=np.sum),)
        .reset_index()
    )

    print(cost_df_pivot.head())

    cost_df_pivot.to_excel(irby_workbook, sheet_name="Audit List", index=False)

    irby_workbook.save()

    #
    # irby_file = "c:\\users\\u6zn\\downloads\\WO_Shipment_Update_20210702_230002.dat"
    # irby_file_df = pd.read_csv(irby_file)
    #
    # print(irby_file_df.dtypes)
    #
    # irby_with_cost = irby_file_df.merge(cost_df, how="left", on="WONUM")
    #
    # print(irby_with_cost.head())
    # irby_with_cost.to_excel(
    #     f"c:\\users\\u6zn\\desktop\\irby_with_cost_{timestring}.xlsx"
    # )
    #
    # velocity_file = "c:\\users\\u6zn\\downloads\\VCTY_monthly_kpi (5).csv"
    # velocity_file_df = pd.read_csv(velocity_file)
    # velocity_with_cost = velocity_file_df.merge(irby_file_df, how="left", on="ITEMNUM")
    # velocity_with_cost.to_excel(
    #     f"c:\\users\\u6zn\\desktop\\velocity_with_cost_{timestring}.xlsx"
    # )


# TODO retrieve Velocity report using Selenium (maybe not since it's monthly)
# TODO summarize cost per WO using the WPMATERIAL report
# TODO calculate which WO to audit based on 20% of dollar value and LANE status

if __name__ == "__main__":
    main()
