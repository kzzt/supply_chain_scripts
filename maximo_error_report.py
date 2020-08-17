# queries for inventory errors- negative curbal, avgcost, or reorder/vmi mixups

import win32com.client as win32
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

sql_file = open(cfg.inv_errors, "r")
inv_errors_query = sql_file.read()
sql_file.close()


def report_emailer(df_report):
    outlook = win32.Dispatch("outlook.application")
    recipient = "kevin.vaughan2@oncor.com"
    subject = "Inventory Errors in Maximo"
    mail = outlook.CreateItem(0)
    mail.To = recipient
    # mail.CC =
    mail.Subject = subject
    # mail.Attachments.Add(file)
    inv_err_html = f"""<h3>The following items need attention.</h3>
                       {df_report.to_html()}"""
    mail.HtmlBody = inv_err_html
    mail.Send() ## uncomment or it won't work
    # mail.Display(False)


def main():
    inv_err_df = pd.DataFrame(pd.read_sql_query(inv_errors_query, engine))
    inv_err_df.columns = map(str.upper, inv_err_df.columns)

    if len(inv_err_df.index) > 0:
        print(inv_err_df)

        print("Mail sent!")
        return report_emailer(inv_err_df)
    else:
        print("No action needed.")


if __name__ == "__main__":
    main()
