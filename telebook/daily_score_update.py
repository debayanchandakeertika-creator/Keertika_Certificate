import requests
import time
import datetime
from datetime import timedelta
import sys


current_time=datetime.datetime.now().date()
current_hour=datetime.datetime.now()

print(current_time.day)


if current_hour.hour==8:
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Kolkata&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Kolkata-112&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Kolkata-113&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Bhubaneswar&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Durgapur&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Assam&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Mumbai&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Siliguri&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=Gujarat&days=1")
	time.sleep(5)
	score_kol=requests.get("http://188.166.10.173:1995/calculate_score?branch=IFM-KOL&days=1")
	time.sleep(5)
	score_kol1=requests.get("http://188.166.10.173:1995/backup_data?mode=delete_succ")
	time.sleep(5)

	





