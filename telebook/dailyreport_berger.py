import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==6:
	# send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Berger Paints")
	send_mail=requests.get("http://188.166.10.173:1995/qrReporthouse?sitecode=AMRI Hospital BBR")	
	time.sleep(10)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=MK Enterprise")
	time.sleep(5)

if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Karuna Management Services")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Apeejay House")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Maa Laxmi RELS 01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Maa Laxmi LINS 02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=South Wind Project")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Associated Minerals")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Altamira Project")
	time.sleep(5)

if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=YOKOGAWA INDIA LIMITED")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Ganesh Grains Ltd")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=CFS PHARMA AND OPTICALS")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=NEW DELHI CENTER FOR SIGHT")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Tirumala Balaji Alloys")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Shree Mangal Trust")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Merlin Projects Ltd")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Calcutta Lions Netra Niketan CLSN00001")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Ekkya House E006 11")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Emami House E006 03")
	time.sleep(5)






