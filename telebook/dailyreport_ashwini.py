import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==9:
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu?sitecode=Ashwini Hospital")

	send_mail=requests.get("http://188.166.10.173:1995/qrReportpostpatrol?sitecode=The SOUL PCSA02")
	# print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=India Auto Mobile")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Emami Foundation")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Randstad India M012 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=ACC CEMENT ACCL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Pioneer Packaging PPCI02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=EMAMI REALTY LTD EIL201")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=PS PACE OWNERS ASSOCIATIONS")
	time.sleep(5)