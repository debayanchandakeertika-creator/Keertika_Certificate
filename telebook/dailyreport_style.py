import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Style Baazar")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=South Calcutta Girls College")
	time.sleep(5)
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=YORK PRINT YPRT02")
	# time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Apollo Hospitals")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Aum Packaging APKG01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Piramal Mahalakshmi")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Sri Siromani Dealers SRSI01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=M J INDUSTRIAL PARK MITP01")
	time.sleep(5)
