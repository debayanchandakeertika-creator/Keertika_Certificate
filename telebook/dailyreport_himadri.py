import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==10:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Himadri Chemical Ltd.")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Calcutta Television Network CLTN00001")
	print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Calcutta Television Network CLTN00002")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Jal Pump J001 02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Anjali Jewellers ANJL 03")
	time.sleep(5)