import requests
import time
import datetime
from datetime import timedelta
# import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==10:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Teloijan-Tea")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=South Asian Red FM SHAL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Astral Flat Owners ASAF02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=MERLIN RESIDENCY APPARTMENT MRAA01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Great Eastern Stores APHL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Kharagpur Metal Reforming Industries KRRI02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Masterly Kolkata Facility Maintenance CLIS02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Kharagpur Metal Reforming Industries KRRI03")
	time.sleep(5)