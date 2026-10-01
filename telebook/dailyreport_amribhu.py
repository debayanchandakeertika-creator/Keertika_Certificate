import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu?sitecode=AMRI-Bhubaneswar")
	time.sleep(5)