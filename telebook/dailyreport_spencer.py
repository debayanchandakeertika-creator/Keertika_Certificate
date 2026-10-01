import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==6:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=East India Pharmaceutical Works EIPH01")
	# print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Aditya Birla AFRL02")
	# print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportabc?sitecode=Apeejay School APJY01")
	# print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL82")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL63")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Silverleaf Tea SRLT01")

if current_hour.hour==7:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL05")
	# print(send_mail)
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL06")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL07")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Spencer Retail SPRL67")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Wow Momo Foods SIL MOMO01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=ACC Limited SIL ACCL01")
	time.sleep(5)

if current_hour.hour==8:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Ganesh Grains GHGS04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Nemichand Bamalwa Jewellers NMCJ01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Emami Limited GUW EMMI06")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Priyanka Agencies EMMI01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Aditya Birla AFRL04")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Aditya Birla AFRL08")

if current_hour.hour==9:
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=ACC Limited ACCL01")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=ACC Limited ACCL02")
	time.sleep(5)
	# send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Hathway Digital HWCC01")
	# time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReportbhu2s?sitecode=Assam Hospital AMHL03")
	time.sleep(5)







