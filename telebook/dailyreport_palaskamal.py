import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==10:

	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=UDAYAN ENTREPRENEURS WELFARE UNWA01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Plumbum Chemicals and Pigments PLCA01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=PS PRIMARC PROJECTS LLP PRMP03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Ganesh Grains GHGS 00002")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Vishva Electrotech Ltd VELI02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=PAR Enterprises PRES 0001")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=AGS Retails Ind Ltd AGSL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Fortune Marketing FMTG01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=SE Builders And Realtors UKAC02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=SE Builders And Realtors UKAC01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Sonotel Hotel and Resort OZNE04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Parklane Hotel Resorts RNHL01")
	time.sleep(5)


	
