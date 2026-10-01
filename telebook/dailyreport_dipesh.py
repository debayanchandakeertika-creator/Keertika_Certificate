import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==11:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Maa Laxmi LINS 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Super Forging and Steel SSTS01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Calcutta Club Limited CACL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Sanjay Creation SCIN 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=City Bazaar Metro CIBM01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Bazaar Kolkata BZKT 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Aspira Loharuka Developers BGBN08")
	time.sleep(5)