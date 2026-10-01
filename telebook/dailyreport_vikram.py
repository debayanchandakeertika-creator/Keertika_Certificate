import requests
import time
import datetime
from datetime import timedelta
# import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Vikram Solar")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabg?sitecode=Colliers International Property Services CLIS01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Anjali Jewellers ANJL 04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Omkareshwar Mahadev Trust GSAE01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Mega Mall APAI01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Wellside Developers WLLD01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Eco Jute EJUE02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Maxcab Wire and Cables MWSL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Coastal Roadways CLLD01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Twenty One Century TOCP01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Manaksia Limited MNKS01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Spencer Retail SPRL10")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Sundarban Livestock SCEM01")
	time.sleep(5)