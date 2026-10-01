import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==10:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Pearson Drums and Barrels")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=All Idia Trinomul Congress ATMC01")
	print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Urvashi Apartments UIAO01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Spencer Retail SPRL49")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL26")
	time.sleep(5)



if current_hour.hour==11:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Abhyagama Hospitality LLP ABHY01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Saltee Spacio Owners Association Ltd SOAL01")
	print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Natural City Flat Owners NLWA01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Tarsons Products Ltd TRSS03")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Meghbela")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Manjushree Mlantation MNJS02")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=EMAMI REALTY LTD ERYD01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Jal Vayu Towers JVAT01")
	time.sleep(5)
