import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()


if current_hour.hour==11:
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Hiland Sapphire")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Orbit Tower")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=VOLTAMP Electrical")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Narayan Hospital")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Rameswaram Mansion")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=Veekay Industries")
	time.sleep(5)
	send_mail=requests.get("http://188.166.10.173:1995/qrReport2s?sitecode=MK MKEE 03")

