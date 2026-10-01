import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

print(current_time.day)


if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportevent?sitecode=Kolkata International Poultry Fair&atime=8&btime=20&ctime=0&freq=60&night=0")
	time.sleep(5)

	# if current_time.day==25 and current_time.month==1:
	# 	send_mail=requests.get("http://188.166.10.173:1995/qrReportevent?sitecode=PC Chandra Garden&atime=9&btime=23&ctime=0&freq=60&night=0")
	# 	time.sleep(5)
	# 	# send_mail=requests.get("http://188.166.10.173:1995/qrReportevent?sitecode=Eastern Metropolitan Club&atime=1&btime=17&ctime=0&freq=60&night=0&removea=1")
	# 	# time.sleep(5)




