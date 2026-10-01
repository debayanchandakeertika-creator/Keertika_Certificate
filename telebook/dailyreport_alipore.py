import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==11:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Oriental Cusien")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Global Motocorp LLP")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=WOW MOMO KASBA 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=WOW MOMO KASBA 02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Anamika Kala Sangam Trust")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Aspira Loharuka Developers BGBN08")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=4 Sight Model Town")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=CRI Ltd CRIL02")
	time.sleep(5)