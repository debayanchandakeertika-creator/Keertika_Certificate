import requests
import time
import datetime
from datetime import timedelta
# import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

if current_hour.hour==10:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Dalmia Tea Plantation TPIL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Dalmia Tea Plantation TPIL02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Dalmia Tea Plantation TPIL03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Hathway Digital HWCC01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Desun Hospital and Heart Institute HLAI01")
	time.sleep(5)
