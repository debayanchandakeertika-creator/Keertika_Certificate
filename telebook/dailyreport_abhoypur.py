import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==10:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportE?sitecode=Emami-Abhoypur")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Goethe Institute GTHI 01")
	time.sleep(5)