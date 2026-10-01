import requests
import time
import datetime
from datetime import timedelta


current_time=datetime.datetime.now()

current_time1=current_time-timedelta(days=1)

if current_time.hour == 11:
	startdate=current_time1.strftime("%Y-%m-%d")
	enddate=current_time1.strftime("%Y-%m-%d")
	print(startdate,enddate)

	yo=requests.get(f'http://188.166.10.173:7878/airportreport?site=NSB International Airport&startdate={startdate}&enddate={enddate}')

	print(yo)





# http://188.166.10.173:7878/airportreport?site=NSB International Airport&startdate=2023-04-19&enddate=2023-04-19