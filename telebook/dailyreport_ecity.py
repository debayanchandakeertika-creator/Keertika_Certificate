import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==9:
	#send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=E-CITY Mall")
	# time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=AMRI Hospitals A01701")
	time.sleep(5)