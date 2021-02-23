# queries VMI po from the previous week and outputs to file named for PONUMs
import cx_Oracle

import pandas as pd
from sqlalchemy import create_engine

import onccfg as cfg

user = cfg.username
password = cfg.password
sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

cstr = f"oracle://{user}:{password}@{sid}"

engine = create_engine(cstr, convert_unicode=False, pool_recycle=10, pool_size=50, echo=False)

sql_file = open(cfg.vmi_query, "r")
vmi_query = sql_file.read()
sql_file.close()



def main():
    vmi_df = pd.DataFrame(pd.read_sql_query(vmi_query, engine))
    vmi_df.columns = map(str.upper, vmi_df.columns)

    vmi_df.loc[
        vmi_df["DAYS STOCK INCLUDING ON ORDER"].lt(0) | vmi_df["DAYS STOCK INCLUDING ON ORDER"].isnull(),
        "DAYS STOCK INCLUDING ON ORDER",
    ] = 0

    filename_str = " & ".join(map(str, vmi_df["PONUM"].unique()))

    print(filename_str)

    writer = pd.ExcelWriter(f"Z:\\SHAWN_MORSE\\VMI PO Validation\\{filename_str}.xlsx")

    vmi_df.to_excel(writer, "VMI_PO", index=False)
    workbook = writer.book
    worksheet = writer.sheets["VMI_PO"]
    worksheet.set_column("F:F", 32)

    writer.save()

    timeFix(filename_str)


if __name__ == "__main__":
    main()
