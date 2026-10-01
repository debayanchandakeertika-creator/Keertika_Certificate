import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport?sitecode=AMRI-Saltlake")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=BD Memorial School BDML04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=BD Memorial School BDML05")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Chandrani Enterprises KOL CNDR01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Chandrani Enterprises CNDR03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Chandrani Enterprises KOL CNDR02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Pulse Diagnostics PLSE03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Chandrani Enterprises CNDR02")
	time.sleep(5)