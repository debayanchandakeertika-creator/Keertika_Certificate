import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Park Hospital")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Sterling Pharmacy")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=TK SENS")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Kharagpur Metal KRRI 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Kharagpur Metal KRRI 02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabcgm?sitecode=YORK PRINT YPRT03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Emami Corporate")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Torsa Machines")
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Vishakha Polyfab")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Ahura Mazda")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabcgm?sitecode=YORK PRINT YPRT04")
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=TARSON")
	# time.sleep(5)
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Ms S N ENTERPRISE")
	# time.sleep(5)
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Artistery House")



