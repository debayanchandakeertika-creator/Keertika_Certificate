import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

print(current_time.day)


if current_hour.hour==11:
	score_kol=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=International Management Institute INIT02")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Piramal Agasthya")
	time.sleep(5)
	# score_kol=requests.get("http://188.166.10.173:1995/qrReportpostpatrol?sitecode=The SOUL PCSA02")
	# time.sleep(5)
