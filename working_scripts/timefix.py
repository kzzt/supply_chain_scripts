import os, sys
import datetime
import time

myDate = "2021-02-22 15:22:32"
print(time.mktime(datetime.datetime.strptime(myDate,
                                             "%Y-%m-%d %H:%M:%S").timetuple()))


filepath = "c:\\users\\u6zn\\desktop\\950 dataloads\\Southwire ROP changes at 950 - 2.22.21.csv"

access_time = time.mktime(datetime.datetime.strptime(myDate,
                                             "%Y-%m-%d %H:%M:%S").timetuple())
modification_time = time.mktime(datetime.datetime.strptime(myDate,
                                             "%Y-%m-%d %H:%M:%S").timetuple())
os.utime(filepath, (access_time, modification_time))
