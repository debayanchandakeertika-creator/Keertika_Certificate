import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

current_day=datetime.datetime.today().weekday()

# print(current_day)

if current_hour.hour==10:
	if current_day!=0:
		send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Aeropict Logistics Solutions ARLG01")
		time.sleep(5)
	if current_day!=2:
		send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Future Merchandising FVRL12")
		time.sleep(5)
