import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

print(current_time.day)


if current_hour.hour==7:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=WILO MATHER AND PLATT PUMPS ACPL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=MACKERTICH CONSULTANCY SERVICES MKRH01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Uma Poly Solutions UPLY01")
	time.sleep(5)
if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Pioneer Packaging Industries")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Kamal Mukherjee")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=SRG Exim SEIM00001")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=J G Chemicals J003 02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Dharampal Sottopal")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Ganesh Grains GHGS 00003")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Tarsons Products TRSS01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=ACC Ltd KOL ACCL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Sundrex Oil Company Ltd SOCL01")
	time.sleep(5)









	