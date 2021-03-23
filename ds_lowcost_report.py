import pandas as pd
import os
import glob


filename = pd.read_csv(
    "C:\\users\\u6zn\\Desktop\\lowcost\\Low_cost.csv", dtype={"LOCATION": object}
).groupby("LOCATION")


filename.apply(
    lambda x: x.to_excel(
        f"c:\\users\\u6zn\\desktop\\lowcost\\Low Cost - {x.name}.xlsx", index=False
    )
)

os.chdir("C:\\users\\u6zn\\Desktop\\lowcost")
filelist = glob.glob("*.xlsx")
for file in filelist:
    df = pd.read_excel(file)
    writer = pd.ExcelWriter(file, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Low Cost")

    workbook = writer.book
    worksheet = writer.sheets["Low Cost"]
    worksheet.set_column("A:B", 10)
    worksheet.set_column("C:C", 8)
    worksheet.set_column("D:E", 4)
    worksheet.set_column("F:F", 9)
    worksheet.set_column("G:H", 14)
    worksheet.set_column("I:I", 100)

    writer.save()
    # df.to_excel(writer, index=False, sheet_name= f{location})


# file.apply(lambda x: x.to_excel("c:\\users\\u6zn\\desktop\\low cost\\Low Cost - {}.xlsx".format(x.name), index=False))


# file2 = file[['ITEMNUM','LOCATION','CURBAL','ROP','EOQ','HARD_RESERVE', 'DESCRIPTION']].sort_values(['LOCATION', 'ITEMNUM'])

# file2.to_csv("c:\\users\\u6zn\\desktop\\output.csv", index=False)
"""
border = Border(left=Side(border_style=None,
                           color='FF000000'),
                 right=Side(border_style=None,
                            color='FF000000'),
                 top=Side(border_style=None,
                          color='FF000000'),
                 bottom=Side(border_style=None,
                             color='FF000000'),
                 diagonal=Side(border_style=None,
                               color='FF000000'),
                 diagonal_direction=0,
                 outline=Side(border_style=None,
                              color='FF000000'),
                 vertical=Side(border_style=None,
                               color='FF000000'),
                 horizontal=Side(border_style=None,
                                color='FF000000')
                )

# for i, x in file2.groupby('LOCATION'):
#     p = os.path.join(os.getcwd(), "lowcost_{}.xlsx".format(i.lower()))
#     x.to_excel(p, index=False).worksheet.set_column('B:D', 20)
    """
