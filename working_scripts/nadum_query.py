# queries VMI po from the previous week and outputs to file named for PONUMs
import cx_Oracle

import time
import pandas as pd
from sqlalchemy import create_engine

import win32com.client as win32

import onccfg as cfg

user = cfg.username
password = cfg.password
sid = cx_Oracle.makedsn(cfg.host, cfg.port, sid=cfg.sid)  # host, port, sid

cstr = f"oracle://{user}:{password}@{sid}"

engine = create_engine(
    cstr,
    convert_unicode=False,
    pool_recycle=20,
    pool_size=100,
    echo=False,
    max_identifier_length=128,
)

sql_file = open(cfg.demand_query, "r")
demand_query = sql_file.read()
sql_file.close()

timestring = time.strftime("%m.%d.%y")


def main():
    demand_df = pd.DataFrame(pd.read_sql_query(demand_query, engine))
    demand_df.columns = map(str.upper, demand_df.columns)

    writer = pd.ExcelWriter(
        f"c:\\users\\u6zn\\desktop\\xlsx\\DEMAND STATUS REPORT {timestring}.xlsx"
    )
    demand_df.to_csv(f"c:\\users\\u6zn\\desktop\\xlsx\\DEMAND STATUS REPORT {timestring}.csv", index=False)


    demand_df.to_excel(writer, "DEMAND STATUS", index=False)
    workbook = writer.book
    worksheet = writer.sheets["DEMAND STATUS"]
    worksheet.set_column("B:B", 32)

    writer.save()


if __name__ == "__main__":
    main()

    outlook = win32.Dispatch("outlook.application")

    mail = outlook.CreateItem(0)
    mail.To = "kevin.vaughan@oncor.com; Nadum.Charles@oncor.com"
    mail.Subject = f"Demand Report {timestring}"
    mail.Attachments.Add(f"c:\\users\\u6zn\\desktop\\xlsx\\DEMAND STATUS REPORT {timestring}.csv")
    mail.HtmlBody = (
        "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=us-ascii\"><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><style><!--\n"
        "/* Font Definitions */\n"
        "@font-face\n"
        "	{font-family:\"Cambria Math\";\n"
        "	panose-1:2 4 5 3 5 4 6 3 2 4;}\n"
        "@font-face\n"
        "	{font-family:Calibri;\n"
        "	panose-1:2 15 5 2 2 2 4 3 2 4;}\n"
        "@font-face\n"
        "	{font-family:Verdana;\n"
        "	panose-1:2 11 6 4 3 5 4 4 2 4;}\n"
        "/* Style Definitions */\n"
        "p.MsoNormal, li.MsoNormal, div.MsoNormal\n"
        "	{margin:0in;\n"
        "	margin-bottom:.0001pt;\n"
        "	font-size:11.0pt;\n"
        "	font-family:\"Calibri\",sans-serif;}\n"
        "a:link, span.MsoHyperlink\n"
        "	{mso-style-priority:99;\n"
        "	color:#0563C1;\n"
        "	text-decoration:underline;}\n"
        "a:visited, span.MsoHyperlinkFollowed\n"
        "	{mso-style-priority:99;\n"
        "	color:#954F72;\n"
        "	text-decoration:underline;}\n"
        "span.EmailStyle17\n"
        "	{mso-style-type:personal;\n"
        "	font-family:\"Calibri\",sans-serif;\n"
        "	color:windowtext;}\n"
        "span.EmailStyle18\n"
        "	{mso-style-type:personal;\n"
        "	font-family:\"Calibri\",sans-serif;\n"
        "	color:#1F497D;}\n"
        "span.EmailStyle19\n"
        "	{mso-style-type:personal-reply;\n"
        "	font-family:\"Calibri\",sans-serif;\n"
        "	color:#1F497D;}\n"
        ".MsoChpDefault\n"
        "	{mso-style-type:export-only;\n"
        "	font-size:10.0pt;}\n"
        "@page WordSection1\n"
        "	{size:8.5in 11.0in;\n"
        "	margin:1.0in 1.0in 1.0in 1.0in;}\n"
        "div.WordSection1\n"
        "	{page:WordSection1;}\n"
        "--></style><!--[if gte mso 9]><xml>\n"
        "<o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" />\n"
        "</xml><![endif]--><!--[if gte mso 9]><xml>\n"
        "<o:shapelayout v:ext=\"edit\">\n"
        "<o:idmap v:ext=\"edit\" data=\"1\" />\n"
        f"</o:shapelayout></xml><![endif]--></head><body lang=EN-US link=\"#0563C1\" vlink=\"#954F72\"><div class=WordSection1><p class=MsoNormal><span style='color:#1F497D'>Good morning,</span><span style='color:#1F497D'><o:p></o:p></span></p><p class=MsoNormal><span style='color:#1F497D'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='color:#1F497D'>{timestring}.<o:p></o:p></span></p><p class=MsoNormal><span style='color:#1F497D'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='color:#1F497D'>Thanks,</span><o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><b><span style='font-family:\"Verdana\",sans-serif;color:#7F7F7F'>Kevin Vaughan<o:p></o:p></span></b></p><p class=MsoNormal style='text-autospace:none'><span style='font-family:\"Verdana\",sans-serif;color:gray'>Material Management Analyst<o:p></o:p></span></p><p class=MsoNormal style='text-autospace:none'><b><span style='font-family:\"Verdana\",sans-serif;color:#717073'>Oncor | Supply Chain Management<o:p></o:p></span></b></p><p class=MsoNormal style='text-autospace:none'><span style='font-family:\"Verdana\",sans-serif;color:#717073;mso-fareast-language:ZH-CN'>15101 Trinity Blvd. <o:p></o:p></span></p><p class=MsoNormal style='text-autospace:none'><span style='font-family:\"Verdana\",sans-serif;color:#717073;mso-fareast-language:ZH-CN'>Fort Worth, TX 76155<o:p></o:p></span></p><p class=MsoNormal style='text-autospace:none'><span style='font-family:\"Verdana\",sans-serif;color:gray'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='font-family:\"Verdana\",sans-serif;color:gray'>Tel: 817.359.2310<o:p></o:p></span></p><p class=MsoNormal><span style='font-family:\"Verdana\",sans-serif;color:#7F7F7F'>Cell: 214.542.6626<o:p></o:p></span></p><p class=MsoNormal><span style='font-family:\"Verdana\",sans-serif'><a href=\"mailto:Kevin.Vaughan@oncor.com\">Kevin.Vaughan@oncor.com</a><span style='color:#2E74B5'><o:p></o:p></span></span></p><p class=MsoNormal><span style='color:#1F497D'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><b><span style='font-family:\"Arial\",sans-serif;color:#2E74B5'>oncor.com<o:p></o:p></span></b></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>")

    ## mail.Send()  ## uncomment or it won't work
    mail.Display(False)
