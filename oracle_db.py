import time

import cx_Oracle

timestring = time.strftime("%Y-%m-%d")


user = "u6zn"
userpwd = "U6_Oncor2020zn"  # Obtain password string from a user prompt or environment variable
# dsnStr = cx_Oracle.makedsn("odmoraprd03.corp.oncor.com", "1521", "ora1")


with cx_Oracle.connect(user, userpwd, "odmoraprd03.corp.oncor.com:1560/AMODSSEP") as connection:
    cursor = connection.cursor()

    for row in cursor.execute("SELECT * FROM MXRADS.ITEM WHERE ROWNUM <= 10"):
        print(row)
        
    connection.commit()

