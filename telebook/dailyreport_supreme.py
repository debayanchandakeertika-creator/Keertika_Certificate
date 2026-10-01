import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportpostpatrol?sitecode=Supreme Industries")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportpostpatrol?sitecode=Supreme Industries SPRM01")
	time.sleep(5)
