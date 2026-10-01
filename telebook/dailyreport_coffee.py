import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

if current_hour.hour==11:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Wedlock Greens Hotel MPDS01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=REASTORY COFFEE HOUSE NAIE01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Primarc Project PRMP04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=URBANA FACILITY MAINTANANCE URFL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=PS RESERVE OWNERS ASSOCIATIONS  PSGP01")
	time.sleep(5)
