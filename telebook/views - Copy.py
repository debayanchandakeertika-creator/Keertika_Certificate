from django.shortcuts import render
from django.shortcuts import render,redirect
from book.settings import *
from django.contrib.auth import logout
from django.http import HttpResponse
from django.core.mail import send_mail
from django.contrib import messages
from django.http import HttpResponseRedirect
from django.contrib.auth.models import User, auth
from django.views.decorators.clickjacking import xframe_options_exempt
from telebook.models import *
from time import strptime
import time
import datetime
from django.db.models import Q
from django.utils import timezone
#from pytz import timezone
from django.db.models import Count
import csv
from django.utils.encoding import smart_str
from datetime import timedelta
# from django.utils import timezone
from openpyxl import load_workbook
import os
import json
from django.core.serializers.json import DjangoJSONEncoder
from django.core.serializers import serialize
from django.core.mail import EmailMessage
from django.conf import settings
from django.views.decorators.csrf import csrf_exempt
import urllib.parse
import requests
import pandas as pd
from django.http import JsonResponse,FileResponse
#from telebook import data_generation
import openpyxl
import uuid
import base64
import qrcode
import img2pdf
from PIL import Image,ImageDraw,ImageFont
from datetime import date, time, datetime
from shutil import copyfile
from django.contrib.auth.forms import AuthenticationForm
from django.contrib import messages
import numpy as np
import math
import openpyxl as xl
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from openpyxl.utils import get_column_letter


# def data_test(request):

# 	tz=timezone.get_current_timezone()

# 	time1=datetime.now()

# 	filter_time=time1-timedelta(days=90,hours=time1.hour,minutes=time1.minute)

# 	filter_time=tz.localize(filter_time)
# 	now_time=tz.localize(time1)



# 	print(now_time,filter_time)

# 	df_all=pd.DataFrame(list(pathTbl.objects.filter(time__range=(filter_time,now_time)).values()))
# 	df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
# 	all_uuid=pd.DataFrame(list(uuidTbl.objects.all().values()))
# 	all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
# 	all_place=pd.DataFrame(list(placeTbl.objects.all().values()))
# 	all_place=all_place.rename(columns={"placeid":"placeid_id"})

# 	uuid_master=pd.merge(all_uuid,all_place,on="placeid_id")
# 	uuid_master=uuid_master.rename(columns={"uuid":"uuid_id"})
# 	print(uuid_master.columns)

# 	all_emp.to_pickle("allemp.pkl")

# 	# all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

# 	# df_master_1=pd.merge(df_all,uuid_master,on="uuid_id")
# 	# df_master=pd.merge(df_master_1,all_emp,on="empid_id")

# 	# print(len(df_master))

# 	# df_master.to_pickle("test1.pkl")

# 	# df_master.to_csv("emp_test_data.csv")



# 	# print(all_uuid.columns)
# 	# print(df_all.columns)
# 	# print(all_emp.columns)
# 	# print(all_place.columns)

# 	# print(len(df_path1))
# 	return JsonResponse({"result":"success"})



# def get_all_emp(request):

# 	branch=request.GET.get("branch","")

# 	if branch=="Assam":
# 		sitelist=list(siteTbl.objects.filter(Q(location=branch)|Q(location="Guwahati")).values_list('name',flat=True))


# 	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__in=sitelist).values()))


# 	total_list=all_path["empid_id"].unique()

# 	all_emp=pd.DataFrame(list(empTbl.objects.filter(empid__in=total_list).values()))

# 	print(all_emp)

# 	all_emp.to_csv("allemp1.csv")


# 	return HttpResponse("Success")

def version_check(request):
	curr_version=request.GET.get("version","")

	org_version=3

	if curr_version=="":
		return JsonResponse({"status":"notok"})
	elif (org_version-float(curr_version))>1:
		return JsonResponse({"status":"notok"})
	else:
		return JsonResponse({"status":"ok"})

def backup_data(request):
	mode1=request.GET.get("mode","")
	authkey=request.GET.get("auth","")


	if mode1=="":
		return JsonResponse({"status":"failed"})
	elif mode1=="delete":
		if authkey!="xGbp5648Lt":
			return JsonResponse({"status":"failed"})
		else:
			tz=timezone.get_current_timezone()

			time1=datetime.now()

			filter_time=time1-timedelta(days=120,hours=time1.hour,minutes=time1.minute)

			filter_time=tz.localize(filter_time)



			print(filter_time)

			df_path1=pd.DataFrame(list(pathTbl.objects.filter(time__lte=filter_time).values()))

			print(len(df_path1))

			df_path1.to_csv("backup_path.csv")

			dt_1=filter_time.strftime("%d %B %Y")
			print(dt_1)

			mail_sender=False

			try:
				# mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
				# df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

				# # print(df)

				# df_1=df.loc[df['siteid']==siteid]
				# mail_list=df_1.iloc[:,1:].values.tolist()
				# mail_list=[element for element in mail_list[0] if '@' in str(element)]
				# send_to_list=mail_list
				#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
				send_to_list=['arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
				subject=f"Backup Database"
				try:
					body=f"Dear Sir,\n\nPlease find attached backup of pathTbl of our database for the date less than and equal to {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
				except:
					body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
				# html_message = render_to_string('table/table.html')
				# text_content=strip_tags(html_message)
				# message = get_template('table/table_1.html').render()
				# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
				# email.content_subtype ="html"# Main content is now text/html

			# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
			# 	# email.attach_alternative(html_message,"text/html")

				mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
				mail.attach_file("backup_path.csv")
				mail.send()
				mail_sender=True
			# print("done!")
			# return HttpResponse("Hello World", content_type="text/plain")
			except:

				return JsonResponse({"status":"mail sending failed. Please check server to get the backup csv file"})


				# send_to_list=['arnab@achillesresolute.com']
				# body=body=f"Dear Sir,\n\nBack sending failed, backup is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
				# mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
				# mail.send()

			# module_dir = os.path.dirname(__file__)   #get current directory
			# file_path3 = os.path.join(module_dir,'database.csv')
			if mail_sender==True:
				pathTbl.objects.filter(time__lte=filter_time).delete()
				return JsonResponse({"status":"mail sending successfull. Data is deleted."})
			else:
				return JsonResponse({"status":"mail sending not successfull. Data is not deleted."})



		return JsonResponse({"status":"success"})

	elif mode1=="delete_succ":

		time1=datetime.now()

		filter_time=time1-timedelta(days=26)

		filter_time=filter_time.date()

		print(filter_time)

		successTbl.objects.filter(date1__lte=filter_time).delete()
		# print(len(y))

		return JsonResponse({"status":"success"})





	else:


		return JsonResponse({"status":"success"})

def privacypolicy(request):

	return render(request, 'privacy.html')





@csrf_exempt
def add_post_order(request):
	if request.method=="GET":
		all_site_list=list(siteTbl.objects.all().values_list('siteid'))
		all_site_list1=[]
		for ele in all_site_list:
			all_site_list1.append(ele[0])
		# print(all_site_list1)
		context={"sites":all_site_list1}
		return render(request, 'selectsite.html',context)

	elif request.method=="POST":
		siteid=request.POST.get("sitename","")

		print(siteid)

		all_beacons_names=list(uuidTbl.objects.filter(placeid__siteid__name=siteid).values_list('uuid','placeid_id','name'))

		context={"beacon":all_beacons_names,"site":siteid}


		# print(all_beacons_names)

		return render(request, 'selectsite1.html',context)


def addqs(request):
	uuid=request.GET.get("uuid","")
	qs=request.GET.get("qsno","")
	site=request.GET.get("site","")

	qs=int(qs)

	beacons_names=list(uuidTbl.objects.filter(uuid=uuid).values_list('placeid_id','name'))


	place_name=beacons_names[0][1]+"--"+beacons_names[0][0]

	print(place_name)

	# qrcount=int(empid)

	# qr_available=qr_generator(qrcount)
	# print(qr_available)

	#return JsonResponse({'result':'QR codes are generated'})

	

	# with open('media/qr-img.png', "rb") as image_file:
	# 	image_data = base64.b64encode(image_file.read()).decode('utf-8')

	# ctx["image"] = image_data

	qs_name=[]
	ans_name=[]
	type_name=[]
	shift_name=[]
	start_time=[]
	end_time=[]
	# type_name=[]
	# shift_name=[]

	for i in range (qs):
		nam="qs_name_"+str(i+1)
		nam1="ans_name_"+str(i+1)
		nam2="type_name_"+str(i+1)
		nam3="shift_"+str(i+1)
		nam4="start_time_"+str(i+1)
		nam5="end_time_"+str(i+1)

		qs_name.append(nam)
		ans_name.append(nam1)
		type_name.append(nam2)
		shift_name.append(nam3)
		start_time.append(nam4)
		end_time.append(nam5)

	ret_dict={"qs":qs_name,"ans":ans_name,"type":type_name,"shift":shift_name,"start":start_time,"end":end_time}

	final_d=pd.DataFrame(ret_dict)
	print(final_d)

	json_records = final_d.reset_index().to_json(orient ='records')
	data = []
	data = json.loads(json_records)



	context={"site":site,"place":place_name,"d":data,"qscount":qs,"uuid":uuid}


	return render(request, 'addqs.html',context)


def submitqs(request):

	uuid=request.GET.get("uuid","")
	qs=request.GET.get("qs_count","")
	site=request.GET.get("site_name","")

	print(uuid,qs,site)

	uuidData=uuidTbl.objects.get(uuid=uuid)

	uudata1=str(uuidData)
	uudata1=uudata1.split(",")[1]
	# print(uuidData)
# 
	qs=int(qs)
	# qs_list=[]

	# right_ans_list=[]
	print("This is qs count", qs)


	for i in range(qs):
		x=request.GET.get("qs_name_"+str(i+1),"")
		y=request.GET.get("ans_name_"+str(i+1),"")
		z=request.GET.get("type_name_"+str(i+1),"")
		sh=request.GET.get("shift_"+str(i+1),"")
		start_time=request.GET.get("start_time_"+str(i+1),"")
		end_time=request.GET.get("end_time_"+str(i+1),"")


		data1=list(qsTbl.objects.filter(uuid=uuidTbl.objects.get(uuid=uuid)).values_list("qsid",))
		qsid1=uudata1+"--"+str(len(data1)+1)
		#print(x,y,z,sh)


		abc=qsTbl(qsid=qsid1,uuid=uuidData,qs=x,rightans=y,type1=sh,type2=start_time,type3=end_time).save()




	all_beacons_names=list(uuidTbl.objects.filter(placeid__siteid__name=site).values_list('uuid','placeid_id','name'))

	context={"beacon":all_beacons_names,"site":site,"msg":"Question Addition Success!"}


	# print(all_beacons_names)

	return render(request, 'selectsite1.html',context)








@csrf_exempt

def check_login(request):
	empid=request.GET.get("empid","")
	deviceid=request.GET.get("deviceid","")
	firebaseid=request.GET.get("firebaseid","")

	if empid=="Arnab3150" or empid=="Ajinkya4150":
		a=empTbl.objects.filter(empid=deviceid).exists()

		if empid=="Arnab3150":

			if a:
				b=notificationTbl.objects.filter(empid1="3150").exists()

				if b:
					return JsonResponse({"result": "success","name":"Arnab"})
				else:
					c=notificationTbl.objects.create(empid1="3150",firebaseid=firebaseid,deviceid=deviceid,desg="SE",type1="Arnab",type2="8697037702")
					return JsonResponse({"result": "success","name":"Arnab"})

			else:
				emp=empTbl.objects.create(empid=deviceid,name="Arnab",phone="8697037702")

				b=notificationTbl.objects.filter(empid1="3150").exists()

				if b:
					return JsonResponse({"result": "success","name":"Arnab"})
				else:
					c=notificationTbl.objects.create(empid1="3150",firebaseid=firebaseid,deviceid=deviceid,desg="SE",type1="Arnab",type2="8697037702")
					return JsonResponse({"result": "success","name":"Arnab"})

		elif empid=="Ajinkya4150":
			if a:
				b=notificationTbl.objects.filter(empid1="4150").exists()

				if b:
					return JsonResponse({"result": "success","name":"Ajinkya"})
				else:
					c=notificationTbl.objects.create(empid1="4150",firebaseid=firebaseid,deviceid=deviceid,desg="SE",type1="Ajinkya",type2="8697037702")
					return JsonResponse({"result": "success","name":"Ajinkya"})

			else:
				emp=empTbl.objects.create(empid=deviceid,name="Arnab",phone="8697037702")

				b=notificationTbl.objects.filter(empid1="4150").exists()

				if b:
					return JsonResponse({"result": "success","name":"Ajinkya"})
				else:
					c=notificationTbl.objects.create(empid1="3150",firebaseid=firebaseid,deviceid=deviceid,desg="SE",type1="Arnab",type2="8697037702")
					return JsonResponse({"result": "success","name":"Ajinkya"})

		else:
			pass






	if empid!="" and empid[0]!="0":
		empid="0"+empid

	url=f"http://outsidevendorservice.starspayroll.in:700/CommonServices_OutsideModule/Service_CommonServices.svc/Get_Employee?empid={empid}&emptype=%&empstatus=%&regionid=%&branchid=%&companyid=%&groupcompanyid=%&divisioncode=%"

	try:
		a=requests.get(url)
		print(a)
	except Exception as e:
		print(e)
	data=json.loads(a.text)

	# print(data["Get_EmployeeResult"][0]["userstatus"])

	if len(data["Get_EmployeeResult"])==0:
		return JsonResponse({"result": "Not Found"})

	elif data["Get_EmployeeResult"][0]["userstatus"]!="POSTED":
		return JsonResponse({"result": "Not Found"})
	else:

		a=empTbl.objects.filter(empid=deviceid).exists()

		if a==True:
			empid_org=data["Get_EmployeeResult"][0]["employeeid"]
			empname=data["Get_EmployeeResult"][0]["employeename"]

			b=notificationTbl.objects.filter(empid1=empid_org).exists()
			if b:
				return JsonResponse({"result": "success","name":empname})
			else:
				empname=data["Get_EmployeeResult"][0]["employeename"]
				phone_num=data["Get_EmployeeResult"][0]["cmobileno"]
				desg=data["Get_EmployeeResult"][0]["designation"]

				c=notificationTbl.objects.create(empid1=empid_org,firebaseid=firebaseid,deviceid=deviceid,desg=desg,type1=empname,type2=phone_num)

				return JsonResponse({"result": "success","name":empname})

		else:
			empname=data["Get_EmployeeResult"][0]["employeename"]
			phone_num=data["Get_EmployeeResult"][0]["cmobileno"]
			desg=data["Get_EmployeeResult"][0]["designation"]
			empid_org=data["Get_EmployeeResult"][0]["employeeid"]
			emp=empTbl.objects.create(empid=deviceid,name=empname,phone=phone_num)
			b=notificationTbl.objects.filter(empid1=empid_org).exists()
			if b:
				return JsonResponse({"result": "success","name":empname})
			else:
				empname=data["Get_EmployeeResult"][0]["employeename"]
				phone_num=data["Get_EmployeeResult"][0]["cmobileno"]
				desg=data["Get_EmployeeResult"][0]["designation"]

				c=notificationTbl.objects.create(empid1=empid_org,firebaseid=firebaseid,deviceid=deviceid,desg=desg,type1=empname,type2=phone_num)

				return JsonResponse({"result": "success","name":empname})




			return JsonResponse({"result": "success"})



	return HttpResponse("Success")



def get_emp_data(request):
	empid=request.GET.get("empid","")

	a=empTbl.objects.filter(empid=empid).exists()

	if a:
		b=empTbl.objects.filter(empid=empid).values()
		# print(b[0]['name'])

		return JsonResponse({"status":"Success","name":b[0]['name'],"phone":b[0]['phone']})

	else:
		return JsonResponse({"status":"fail"})










@csrf_exempt
def register(request):
	if request.method=="POST":
		empid=request.POST.get("empid","")
		password1=request.POST.get("pass1","")
		password2=request.POST.get("pass2","")
		mailid=request.POST.get("mailid","")
		desg=request.POST.get("desg","")

		user=User.objects.create_user(username=empid,password=password1,email=mailid)
		user.save()

		return JsonResponse({'result':'Success!'})



	else:

		return render(request, 'register.html')

@csrf_exempt
def login(request):

	if request.method=="POST":
		empid=request.POST.get("empid","")
		password1=request.POST.get("pass1","")
		desg=request.POST.get("desg","")
		print(empid,password1)
		user=auth.authenticate(username=empid,password=password1)
		print(user)

		if user is not None:
			auth.login(request,user)
			return redirect("/addnewsite")
		else:
			messages.info(request,"invalid credentials")
			return redirect("/login")

	else:

		return render(request, 'login.html')
	
@csrf_exempt
def sendqrmail(request):
	if request.method=='POST':
		site_name=request.POST.get("site_name","")
		mail_id=request.POST.get("mail_id","")
		print(site_name)

		path_contain=os.listdir(f"QR_CODES/{site_name}/")
		print(f"QR_CODES/{site_name}/")
		print(path_contain)

		# # try:

		body=f"Dear sir,\n\nPlease find the attached generated QR codes for the site id : {site_name}.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		send_to_list=[mail_id]
		mail = EmailMessage(f"QR CODES | site id: {site_name}",body, settings.EMAIL_HOST_USER, send_to_list)
		for ele in path_contain:
			mail.attach_file(f"QR_CODES\\{site_name}\\{ele}")
		mail.send()
		return JsonResponse({'result':'Success!'})


def qr_pdf_generator(uuid_code,place,site,floor,building,pl_id,typ,shft):

	placeid=placeTbl.objects.get(siteid__name__icontains=site,floor__icontains=floor,building__icontains=building,placeid=pl_id)


	#print(placeid)
	if typ=="patrol":
		abc= uuidTbl(uuid=uuid_code,placeid=placeid,name=place,other=shft,uuidtype='QR PATROL',time1=datetime.now()).save()
	elif typ=="house":
		abc= uuidTbl(uuid=uuid_code,placeid=placeid,name=place,other=shft,uuidtype='HOUSEKEEPING',time1=datetime.now()).save()
	else:
		abc= uuidTbl(uuid=uuid_code,placeid=placeid,name=place,other=shft,uuidtype='QR',time1=datetime.now()).save()


	qr_img = qrcode.make(uuid_code,box_size=30,border=4,)
	qr_img.save(f"QR_CODES/{uuid_code}.png")

	img1 = Image.open(f"QR_CODES/{uuid_code}.png")
	img2 = Image.open("blank1.png")

	Image.Image.paste(img2, img1, (17, 160))
	#img2.save("qr-img.png")

	draw = ImageDraw.Draw(img2)

	draw.text((300, 20),"NIS Management Limited",'black',font = ImageFont.truetype("arial.ttf",40)) 
	draw.text((350, 65), floor + " : "+place+"\n", 'black', font = ImageFont.truetype("arial.ttf",30)) 
	draw.text((300, 95), "Building : " + building+"\n", 'black', font = ImageFont.truetype("arial.ttf",30)) 
	draw.text((250, 125), "Site : " + site+'\n', 'black', font = ImageFont.truetype("arial.ttf",30))
  
	if typ=="patrol": 
		draw.text((450, 160),"(Patrol Monitoring)",'black',font = ImageFont.truetype("arial.ttf",20)) 
	elif typ=="house":
		draw.text((520,150 ),"(Housekeeping)",'black',font = ImageFont.truetype("arial.ttf",30))	
	else: 
		draw.text((450,160),"(Post Monitoring)",'black',font = ImageFont.truetype("arial.ttf",20))



	# draw.text((380, 20),"NIS Management Limited",'black',font = ImageFont.truetype("arial.ttf",45))


	# draw.text((380, 68),floor+": "+place,'black',font = ImageFont.truetype("arial.ttf",40))
	# draw.text((300, 108),building+ ", Site: "+site,'black',font = ImageFont.truetype("arial.ttf",40))
	# if typ=="patrol":
	# 	draw.text((520,150 ),"(Patrol Monitoring)",'black',font = ImageFont.truetype("arial.ttf",30))
	# elif typ=="house":
	# 	draw.text((520,150 ),"(Housekeeping)",'black',font = ImageFont.truetype("arial.ttf",30))
	# else:
	# 	draw.text((520,150 ),"(Post Monitoring)",'black',font = ImageFont.truetype("arial.ttf",30))


	img2.save(f"QR_CODES/{uuid_code}.png")



	pdf_bytes = img2pdf.convert(f"QR_CODES/{uuid_code}.png")
	pdf_path=f"QR_CODES/{uuid_code}.pdf"
	print(f"QR_CODES/{uuid_code}.pdf")
	file = open(pdf_path, "wb")
	file.write(pdf_bytes)
	file.close()

	
	dest_path=f"QR_CODES/{site}/{place}.pdf"
	copyfile(pdf_path,dest_path)

	os.remove(f"QR_CODES/{uuid_code}.png")

def qr_generator(qr_num):
	qr_list=[]
	k=0

	while k<qr_num:
		gen_qr=str(uuid.uuid4())
		if gen_qr not in qr_list:
			qr_list.append(gen_qr)
			k=k+1

	return qr_list

@csrf_exempt
def pdf_gen(request):
	if request.method == 'POST':
		count1=request.POST.get("qr_con","")
		site_name=request.POST.get("site_name","")
		place_list=[]
		uuid_list=[]
		floor_list=[]
		build_list=[]
		plid_list=[]
		type_list=[]
		shift_list=[]
		#try:
		count1=int(count1)
		
		for ele in range (count1):
			x=request.POST.get("qr_place_"+str(ele+1),"")
			y=request.POST.get("place_fl_"+str(ele+1),"")
			z=request.POST.get("build_name_"+str(ele+1),"")
			typ=request.POST.get("type_"+str(ele+1),"")
			shft=request.POST.get("shift_"+str(ele+1),"")

			pl_id=z+"-"+y+"("+site_name+")"

			siteid=siteTbl.objects.get(name__icontains=site_name)
			# print(placeid)
			placeTbl(placeid=pl_id,name=pl_id,siteid=siteid,floor=y,building=z).save()

			place_list.append(x)
			floor_list.append(y)
			build_list.append(z)
			plid_list.append(pl_id)
			type_list.append(typ)
			shift_list.append(shft)
		print(place_list)

		for ele1 in range(count1):
			uuid_x=request.POST.get("uuid_val_"+str(ele1+1),"")
			uuid_list.append(uuid_x)
		print(uuid_list)

		for idx, val in enumerate(uuid_list):
			pl_n=place_list[idx]
			qr_pdf_generator(val,pl_n,site_name,floor_list[idx],build_list[idx],plid_list[idx],type_list[idx],shift_list[idx])
			print(ele)
		download_dict={'place':place_list,'uuid':uuid_list}

		pd_d=pd.DataFrame(download_dict)

		json_records = pd_d.reset_index().to_json(orient ='records')
		data = []
		data = json.loads(json_records)

		with open('media/qr-img.png', "rb") as image_file:
			image_data = base64.b64encode(image_file.read()).decode('utf-8')
		context = {'d': data,'img':image_data,'site':site_name}

		return render(request, 'index2.html', context)


			#return JsonResponse({'result':'Success!'})
		#except:
			#return JsonResponse({'result':'Failed!'})
	


@csrf_exempt
def qrgeneration(request):

	# token=request.GET.get("token","")
	# ttk='FRSvnoKJnlnloaenvhfl74avdw1ASD14av6w1631tf6A8h1thf61fgYhd41g39d1Lgawa1e3aEe3K32asNdsa'
	# if ttk != token:
	# 	return JsonResponse({'result':f'Invalid token'})
	# empid=request.GET.get("empid","")
	# print(os.getcwd(),empid)
	if request.method == 'POST':
		empid=request.POST.get("empid","")
		sitename=request.POST.get("siteid","")

		build_n=request.POST.get("build_name","")
		fl_n=request.POST.get("fl_name","")

		if empid=="":
			return JsonResponse({'result':'Please enter a valid number'})

		else:

			qrcount=int(empid)

			qr_available=qr_generator(qrcount)
			print(qr_available)

			#return JsonResponse({'result':'QR codes are generated'})

			

			with open('media/qr-img.png', "rb") as image_file:
				image_data = base64.b64encode(image_file.read()).decode('utf-8')

			# ctx["image"] = image_data

			filed_name=[]
			uuid_name=[]
			floor_name=[]
			building_name=[]
			type_name=[]
			shift_name=[]

			for i in range (qrcount):
				nam="qr_place_"+str(i+1)
				nam1="uuid_val_"+str(i+1)
				nam2="place_fl_"+str(i+1)
				nam3="build_name_"+str(i+1)
				nam4="type_"+str(i+1)
				nam5="shift_"+str(i+1)
				filed_name.append(nam)
				uuid_name.append(nam1)
				floor_name.append(nam2)
				building_name.append(nam3)
				type_name.append(nam4)
				shift_name.append(nam5)


			ret_dict={}
			ret_dict["qr_code"]=qr_available
			ret_dict['field_nam']=filed_name
			ret_dict['uuid_nam']=uuid_name
			ret_dict['floor_nam']=floor_name
			ret_dict['build_nam']=building_name
			ret_dict["type_name"]=type_name
			ret_dict["shift_name"]=shift_name
			final_d=pd.DataFrame(ret_dict)

			json_records = final_d.reset_index().to_json(orient ='records')
			data = []
			data = json.loads(json_records)

			if build_n!="" and fl_n!="":
				context = {'d': data,'img':image_data,'con':qrcount,'site':sitename,'build':build_n,'fl':fl_n}

			else:
				context = {'d': data,'img':image_data,'con':qrcount,'site':sitename,'build':build_n,'fl':fl_n}



			# data = []
   #  		data = json.loads(json_records)
			return render(request, 'index1.html', context)






	#return render(request, 'index.html', {'form': 'abcd'})


def downloadqr(request):
	qr_val=request.GET.get("qrvalue","")
	place_nam=request.GET.get("place","")

	path1=f"QR_CODES/{qr_val}.pdf"

	with open(path1, 'rb') as pdf: 
		response = HttpResponse(pdf.read()) 
		response['content_type'] = 'application/pdf' 
		response['Content-Disposition'] = f'attachment;filename={place_nam}.pdf'
	return response 
@csrf_exempt
def addnewsite(request):
	comp_list=companyTbl.objects.values()
	print(comp_list)
	print(type(comp_list))

	context={'d':comp_list}

	if request.method == 'POST':
		sitename=request.POST.get("sitename","")
		place=request.POST.get("placeid","")
		company_id=request.POST.get("option","")
		method=request.POST.get("method","")
		try:
			os.mkdir(f"QR_CODES/{sitename}")
		except:
			pass

		if method=="2":

			company_id1=companyTbl.objects.get(name__icontains=company_id)

			# company_id=int(company_id)

			siteTbl(companyid=company_id1,siteid=sitename,name=sitename,location=place).save()

			return render(request, 'index.html', {'site':sitename})
		else:
			company_id1=companyTbl.objects.get(name__icontains=company_id)

			# company_id=int(company_id)

			siteTbl(companyid=company_id1,siteid=sitename,name=sitename,location=place).save()

			return render(request, 'index4.html', {'site':sitename})

			return JsonResponse({'result':'Success!'})






	return render(request, 'index3.html', context)
	#return JsonResponse({'result':'Success!'})


# def valid_user(tid):
# 	try:
# 		obj=users.objects.get(tid=tid).mobile
# 		return True
# 	except Exception as e:
# 		obj= 'None'
# 		return False

# def check_user(tid):
# 	try:
# 		obj=users.objects.get(tid=tid).mobile
# 	except Exception as e:
# 		obj= 'None'
# 	return obj

# def book(request):
# 	#strptime('19-Nov-20','%d-%b-%y')
# 	token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
# 	ttk=request.GET.get("token","")
# 	user=request.GET.get("user","")
# 	if ttk!=token:
# 		success={"abc":"False"}
# 		return render(request,"add_iss.php",success)
# 	else:
# 		if(valid_user(user)):
# 			obj=check_user(user)
# 			data = obj
# 		else:
# 			data="None"
# 		context={"data":data}
# 		return render(request,"lll.php",context)

# def book_meeting(request):
# 	token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
# 	ttk=request.GET.get("token","")
# 	tid=request.GET.get("tid","")
# 	data2=request.GET.get("data","")
# 	user=request.GET.get("user","")
# 	#Book%20for%20%3A%204%20hours%20%3A%20%2030-Nov-20%20%3A%2010
# 	if ttk!=token:
# 		success={"abc":"False"}
# 		return render(request,"add_iss.php",success)
# 	else:
# 		if(valid_user(user)):
# 			obj=check_user(user)
# 			# try:
# 				#data=data2.split(':')
# 			start_time=data2.split(':')[3]
# 			date_to_book=data2.split(':')[2]
# 			duration=data2.split(':')[1].split(' ')[1]
# 			data=start_time+":"+date_to_book+":"+duration
# 			#'30-Nov-20'
# 			data_2=datetime.datetime.strptime(str(date_to_book).strip()+""+start_time, '%d-%b-%y %H')
# 			#data+=data_2
# 			now = data_2
# 			# now = now + datetime.timedelta(hours=6, minutes=30)
# 			aferer = now + datetime.timedelta(hours=int(duration))
# 			#data=meetings.objects.all()
# 			# meet=meetings.objects.filter(start__date=now)
# 			# print(now)

# 			# print(aferer-now)
# 			# print('----------')
# 			# print('strt',now,'end',aferer)
# 			# print(type(now))
# 			# print('----------')
# 			# print(meetings.objects.values('start','end'))
# 			# meet=meetings.objects.filter(start__date=now.date()).filter(end__range=(now,aferer))#.filter(end__lte=aferer)
# 			# print(meetings.objects.filter(start__range=(now,aferer)))
# 			if meetings.objects.filter(start__range=(now,aferer)):
# 				# print('strt')
# 				meet=meetings.objects.filter(start__range=(now,aferer))
# 				# print(meetings.objects.filter(start__range=(now,aferer)))
# 			elif meetings.objects.filter(end__range=(now,aferer)):
# 				# print('end')
# 				meet=meetings.objects.filter(end__range=(now,aferer))
# 				# print(meetings.objects.filter(end__range=(now,aferer)))
# 			else:
# 				meet=[]
# 			# meet=meetings.objects.filter(end__range=now)
# 			# print(meet[0].desk.name)
# 			a=deske.objects.all().values_list('name',flat=True)
# 			desks=[]
# 			for yo in a:
# 				desks.append(yo)
# 			# print('Total desk---',desks)
# 			# data+='Total desk---'+str(desks)+"\n"

# 			the_list=[]
# 			data+="\n"+str(now)+"\n"
# 			for i in meet:
# 				if i.desk.name in desks:
# 					desks.remove(i.desk.name)
# 				# print(i.desk.name)
# 				data+=str(i.start)+"\n"
# 			data+=str(desks)+"\n"
# 			# print(desks)
# 			# print(data)
# 				#data=now
# 			# except Exception as e:
# 			# 	data = "failed"
# 			# 	data = e

# 		else:
# 			data="None"
# 		context={"data":data}
# 		return render(request,"signup.php",context)
#def uuid_generation(request):
	

def database_check(request):
	current_time3=datetime.now()
	current_time3=current_time3-datetime.timedelta(minutes=current_time3.minute, hours=current_time3.hour)

	start_time3=current_time3-datetime.timedelta(days=7)
	end_time3=current_time3
	df_path1=pd.DataFrame(list(pathTbl.objects.filter(time__range=(start_time3,end_time3)).values()))

	module_dir = os.path.dirname(__file__)   #get current directory
	file_path3 = os.path.join(module_dir,'database.csv')

	df_path1.to_csv(file_path3)

	try:

		body=f"Dear Sir,\n\nPlease find attached the database for the period of {start_time3.date()} to {end_time3.date()}.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		send_to_list=['arnab@achillesresolute.com','suman@achillesresolute.com']
		mail = EmailMessage("Database Sheet",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(file_path3)
		mail.send()
	except:
		pass




	success={"abc":'Success'}
	return render(request,"add_iss.php",success)

def grtRoundEmp(eid):
	minimum_cover=1
	totalCount=0
	count=0
	stat=False
	# emppaths=pathTbl.objects.filter(empid=eid).filter(time__date=timezone.localdate()).order_by('time')

	emppaths1=pathTbl.objects.filter(empid__empid=eid,time__date=timezone.localdate()).order_by('time')
	emppaths=[]
	buffer=[]
	for emp in emppaths1:
		emppaths.append(emp.uuid.other)
	for i,ele in enumerate(emppaths):
    # print(i,ele)
		if ele not in buffer:
			buffer.append(ele)
			continue
		if ele in buffer:
			if emppaths[i-1] == ele:
				# print(i,ele)
				pass
			else:
				buffer.append(ele)
	# print(buffer)
	emppaths=buffer
	# print(emppaths.all().values_list('uuid','time'))
	for emppath in emppaths:
		cstat=emppath
		print(cstat)
		# print(emppath.uuid.other)
		if cstat=='start':
			if stat:
				count=0
				stat=False
			else:
				stat=True

		if cstat not in ['start','end'] and stat:
			count+=1
		if cstat=='end' and  stat:
			if minimum_cover <=count:
				totalCount+=1
			stat=False
			count=0
	# print(totalCount)
	return totalCount

	# minimum_cover=2
	# totalCount=0
	# count=0
	# stat=False
	# # emppaths=pathTbl.objects.filter(empid=eid).filter(time__date=timezone.localdate()).order_by('time')

	# emppaths1=pathTbl.objects.filter(empid=eid).order_by('time')
	# emppaths=[]
	# buffer=[]
	# for emppath in emppaths:
	# 	cstat=emppath.uuid.other
	# 	# print(emppath.uuid.other)
	# 	if cstat=='start':
	# 		if stat:
	# 			count=0
	# 			stat=False
	# 		else:
	# 			stat=True

	# 	if cstat not in ['start','end'] and stat:
	# 		count+=1
	# 	if cstat=='end' and  stat:
	# 		if minimum_cover <=count:
	# 			totalCount+=1
	# 		stat=False
	# 		count=0
	# print(totalCount)
	# return totalCount


	# empcount
	# count=
def round_done(emp,intv):
	strt_tm=timezone.datetime.now()-timezone.timedelta(hours=int(intv))
	end_tm=timezone.datetime.now()
	minimum_cover=1
	totalCount=0
	count=0
	stat=False
	# emppaths=pathTbl.objects.filter(empid=eid).filter(time__date=timezone.localdate()).order_by('time')

	emppaths1=pathTbl.objects.filter(time__range=(strt_tm,end_tm),empid__empid=emp).order_by('time')
	# print(emppaths1)
	emppaths=[]
	buffer=[]
	for emp in emppaths1:
		emppaths.append(emp.uuid.other)
	for i,ele in enumerate(emppaths):
    # print(i,ele)
		if ele not in buffer:
			buffer.append(ele)
			continue
		if ele in buffer:
			if emppaths[i-1] == ele:
				# print(i,ele)
				pass
			else:
				buffer.append(ele)
	# print(buffer)
	emppaths=buffer
	# print(emppaths.all().values_list('uuid','time'))
	for emppath in emppaths:
		cstat=emppath
		print(cstat)
		# print(emppath.uuid.other)
		if cstat=='start':
			if stat:
				count=0
				stat=False
			else:
				stat=True

		if cstat not in ['start','end'] and stat:
			count+=1
		if cstat=='end' and  stat:
			if minimum_cover <=count:
				totalCount+=1
			stat=False
			count=0
	# print(totalCount)
	if totalCount>0:
		return 'Yes'
	else:
		return 'No'
def zone(request):
	alll=request.GET.get("all","")
	data=[]
	data1=[]

	empids=pathTbl.objects.values_list('empid',flat=True).distinct()
	if alll != '':
		for emp in empids:
			zonedata=pathTbl.objects.filter(empid=empTbl.objects.get(empid=emp)).values('uuid__other').annotate(rounds=Count('uuid'))
			data.append([empTbl.objects.get(empid=emp).name,zonedata])
		records=pathTbl.objects.order_by('empid','uuid')
		for d in records:
			data1.append({'empid':d.empid.name,'zone':d.uuid.other,'time':d.time})
		context={"data":data,'data1':data1}
		return render(request,"lll1.php",context)
	try:
		days=int(request.GET.get("days",""))
		for emp in empids:
			zonedata=pathTbl.objects.filter(empid=empTbl.objects.get(empid=emp),time__gte=timezone.localtime()-timezone.timedelta(days=3)).values('uuid__other').annotate(rounds=Count('uuid'))
			data.append([empTbl.objects.get(empid=emp).name,zonedata])
		records=pathTbl.objects.filter(time__gte=timezone.localtime()-timezone.timedelta(days=3)).order_by('empid','uuid')
		for d in records:
			data1.append({'empid':d.empid.name,'zone':d.uuid.other,'time':d.time})

	except:
		for emp in empids:
			zonedata=pathTbl.objects.filter(empid=empTbl.objects.get(empid=emp),time__date=timezone.localdate()).values('uuid__other').annotate(rounds=Count('uuid'))
			data.append([empTbl.objects.get(empid=emp).name,zonedata])
		records=pathTbl.objects.filter(time__date=timezone.localdate()).order_by('empid','uuid')
		for d in records:
			data1.append({'empid':d.empid.name,'zone':d.uuid.other,'time':d.time})
	context={"data":data,'data1':data1}
	return render(request,"lll1.php",context)
def rounCal(request):
	#strptime('19-Nov-20','%d-%b-%y')
	# token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	# ttk=request.GET.get("token","")
	# user=request.GET.get("user","")
	# if ttk!=token:
	# 	success={"abc":"False"}
	# 	return render(request,"add_iss.php",success)
	# else:
		intvrl=request.GET.get("intvrl","")
		if intvrl=='':
			data=[]
			empid=pathTbl.objects.values_list('empid',flat=True).distinct()
			for emp in empid:
				count=grtRoundEmp(emp)
				data.append({'empid':emp,'rounds':count})
				# print(emp)
			# data="None"
			context={"data":data}
			return render(request,"lll.php",context)
		else:
			data=[]
			empid=pathTbl.objects.values_list('empid',flat=True).distinct()
			for emp in empid:
				status=round_done(emp,intvrl)
				data.append({'empid':emp,'rounds':status})
				# print(emp)
			# data="None"
			context={"data":data}
			return render(request,"lll.php",context)


@csrf_exempt
def storepostorder(request):

	if request.method == 'POST':
		# print("program is in post method first")

		print(request.FILES['file'])
		print(request.FILES['file'].size)
		fsize=request.FILES['file'].size
		if(fsize <= 200000000 and fsize>10):
			a=handle_uploaded_file(request.FILES['file'])
			try:
				# print("program is in post method")
				y=make_postorder(a)
				if y==True:
					# print("data submission done!")
					os.remove(a)
					return JsonResponse({"result":"success"})

				else:
					return JsonResponse({"result":"fail"})

			except Exception as e:
				return JsonResponse({"result":"fail"})

	else:
		return JsonResponse({"result":"fail"})


def submit_post_order(empid,uuid,qs,ans,time,type1):

	
	if "/" in time:
		time=datetime.strptime(time, '%m/%d/%Y %H:%M').strftime('%Y-%m-%d %H:%M')
	else:
		time=datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')


	empData=empTbl.objects.get(empid=empid)
	abc= qsstoreTbl(uuid=uuidTbl.objects.get(uuid=uuid),qs=qs,ans=ans,time=time,empid=empData,type1=type1).save()
	print("program is here 1")

	return True
def make_postorder(file_name):

	df = pd.read_csv(file_name, header=None,names=['a', 'b', 'c','d','e','f'])
	df=df.dropna()
	# df["f"]=pd.to_datetime(df['f'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
	# print(df)
	# print("prog is in make req")



	for index, rows in df.iterrows():

		# print(rows['f'])
		x=submit_post_order(rows['a'],rows['b'],rows['c'],rows['d'],rows['e'],rows['f'])
		print("program is in loop")

		if x==False:
			print("error occur")
			return False

	return True

@csrf_exempt
def storeData(request):
	token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	uuid=request.GET.get("uuid","")
	mac=request.GET.get("mac","")
	dis=request.GET.get("dis","")
	rssi=request.GET.get("rssi","")
	time=request.GET.get("time","")
	lat=request.GET.get("lat","")
	long1=request.GET.get("long","")
	empid=request.GET.get("empid","")
	# print(time)

	# try:
	# time1=datetime.now()
	# time2=datetime.strptime(time, '%d %B %Y %I:%M:%S %p')
	
		# print(time)
	# except Exception as e:
	# 	time=''
	

	if request.method == 'GET':
		if empid=="aa26071ff9bfd09a":
			time=datetime.now().strftime('%Y-%m-%d %H:%M')
		# print("program is here")

		else:
			# time=datetime.now().strftime('%Y-%m-%d %H:%M')
			time=datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
		
		empid=request.GET.get("empid","")
		
		if abs(int(rssi))<100:
			empData=empTbl.objects.get(empid=empid)
			abc= pathTbl(uuid=uuidTbl.objects.get(uuid=mac),mac=mac,dis=dis,rssi=rssi,time=time,empid=empData,lat=lat,long=long1).save()
			#success={"abc":abc}
			data1=list(qsTbl.objects.filter(uuid=uuidTbl.objects.get(uuid=mac)).values_list("qs","type1"))
			if len(data1)==0:
				return JsonResponse({'result':'Success!'})
			else:
				qs_list=[]
				type_list=[]

				for ele in data1:
					qs_list.append(ele[0])
					type_list.append(ele[1])

				return JsonResponse({'result':'Success!',"qscount":len(data1),"qs":qs_list,"type":type_list})
		else:
			return JsonResponse({'result':'failed!'})
			#success={"abc":'Failed'}
	
	elif request.method == 'POST':
		# print("program is in post method first")

		print(request.FILES['file'])
		print(request.FILES['file'].size)
		fsize=request.FILES['file'].size
		if(fsize <= 200000000 and fsize>10):
			a=handle_uploaded_file(request.FILES['file'])
			try:
				# print("program is in post method")
				y=make_req(a)
				if y==True:
					# print("data submission done!")
					os.remove(a)
			except Exception as e:
				#pass
				os.remove(a)
		return JsonResponse({'result':'Success!'})

		success={"abc":'Success'}
	else:
		success={"abc":'failed'}
		return JsonResponse({'result':'failed!'})

	
	#return render(request,"add_iss.php",success)
	return JsonResponse({'result':'Success!'})
	
		# abc = pathTbl.

def submitdata(empid,uuid,mac,dis,rssi,time,lat,long1):
	# print("program is in submitdata")

	if "/" in time:
		time=datetime.strptime(time, '%m/%d/%Y %H:%M').strftime('%Y-%m-%d %H:%M')
	else:
		time=datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
	if abs(int(rssi))<100:
		empData=empTbl.objects.get(empid=empid)
		abc= pathTbl(uuid=uuidTbl.objects.get(uuid=mac),mac=mac,dis=dis,rssi=rssi,time=time,empid=empData,lat=lat,long=long1).save()
		# print("data submitted successfully")
		# success={"abc":abc}
	else:
		success={"abc":'Failed'}

	return True


def make_req(file_name):

	df = pd.read_csv(file_name, header=None,names=['a', 'b', 'c','d','e','f','g','h'])
	df=df.dropna()
	# df["f"]=pd.to_datetime(df['f'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
	# print(df)
	# print("prog is in make req")

	for index, rows in df.iterrows():

		# print(rows['f'])
		x=submitdata(rows['a'],rows['b'],rows['c'],rows['d'],rows['e'],rows['f'],rows['g'],rows['h'],)
		# x=submitdata(rows['a'],uuid,mac,dis,rssi,time,lat,long1)
		# x=sub

	# with open(file_name, mode='r') as csv_file:
	# 	csv_reader = csv.DictReader(csv_file)

	# 	# print(dir(csv_reader))
	# 	for row in csv_reader:
	# 		print("program is here first")
	# 		print(row)
	# 		for i in row.values():
	# 			print("That is length", len(i))
	# 			print("program is here")
	# 			# if len(i==5):
	# 			# 	return True
	# 			# else:
	# 			empid=urllib.parse.quote(i[0])
	# 			uuid=urllib.parse.quote(i[1])
	# 			mac=urllib.parse.quote(i[2])
	# 			dis=urllib.parse.quote(i[3])
	# 			rssi=urllib.parse.quote(i[4])
	# 			time=urllib.parse.quote(i[5]).replace('%20','+')
	# 			lat=urllib.parse.quote(i[6])
	# 			long1=urllib.parse.quote(i[7])

				
					# url=f"http://127.0.0.1:1995/storeData?empid={empid}&uuid={uuid}&mac={mac}&dis={dis}&rssi={rssi}&time={time}&lat={lat}&long={long1}"
					# print(url)
					# a=requests.get(url)
					# print(a.text)
	return True


def handle_uploaded_file(f):
	import time

	name=f"{time.time()}.csv"
	print(name)
	with open(name, 'wb+') as destination:
		for chunk in f.chunks():
			destination.write(chunk)
	return name

def download_csv_data(request):
	alll=request.GET.get("all","")
	if alll != '':
		# response content type
		response = HttpResponse(content_type='text/csv')
		#decide the file name
		response['Content-Disposition'] = 'attachment; filename="Surveillance_Data_All.csv"'

		writer = csv.writer(response, csv.excel)
		response.write(u'\ufeff'.encode('utf8'))

		#write the headers
		writer.writerow([
			smart_str(u"Emp Name"),
			smart_str(u"Time"),
			smart_str(u"Zone"),
		])
		timezone.localtime()
		#get data from database or from text file....
		events=pathTbl.objects.order_by('empid','uuid','time')
		for event in events:
			writer.writerow([
				smart_str(event.empid.name),
				smart_str(timezone.localtime(event.time)),
				smart_str(event.uuid.other),
			])
		return response

	# response content type
	response = HttpResponse(content_type='text/csv')
	#decide the file name
	response['Content-Disposition'] = 'attachment; filename="Surveillance_Data.csv"'

	writer = csv.writer(response, csv.excel)
	response.write(u'\ufeff'.encode('utf8'))

	#write the headers
	writer.writerow([
		smart_str(u"Emp Name"),
		smart_str(u"Time"),
		smart_str(u"Zone"),
	])
	timezone.localtime()
	#get data from database or from text file....
	events=pathTbl.objects.filter(time__date=timezone.localdate()).order_by('empid','uuid','time')
	for event in events:
		writer.writerow([
			smart_str(event.empid.name),
			smart_str(timezone.localtime(event.time)),
			smart_str(event.uuid.other),
		])
	return response

def download_csv_data1(request):
	days=3
	# response content type
	response = HttpResponse(content_type='text/csv')
	#decide the file name
	response['Content-Disposition'] = 'attachment; filename="Surveillance_Datas.csv"'

	writer = csv.writer(response, csv.excel)
	response.write(u'\ufeff'.encode('utf8'))

	#write the headers
	writer.writerow([
		smart_str(u"Emp Name"),
		smart_str(u"Time"),
		smart_str(u"Zone"),
	])
	# timezone.localtime()
	#get data from database or from text file....
	events=pathTbl.objects.filter(time__gte=timezone.localtime()-timezone.timedelta(days=3)).order_by('time','empid','uuid')
	for event in events:
		writer.writerow([
			smart_str(event.empid.name),
			smart_str(timezone.localtime(event.time)),
			smart_str(event.uuid.other),
		])
	return response

def checkEmp(request):
	empid=request.GET.get("empid","")
	# emp=empTbl.objects.get(empid=empid)
	# print(emp)
	try:
		emp=empTbl.objects.get(empid=empid)
		data='yes,'+empTbl.objects.get(empid=empid).name
	except empTbl.DoesNotExist:
		data='no'
	# if emp:
	# 	data='yes'
	# else:
	# 	data='no'
	context={"data":data}
	return render(request,"lll.php",context)

def get_graph(request):
	days=request.GET.get("days","0")
	try:
		days=int(days)
	except Exception as e:
		days=0
	# print(days)

	prev=days-1
	nextt=days+1

	data=[]
	timed=[x for x in range(0,24)]

	graphDate=timezone.localdate() + timedelta(days=days)

	dayData=pathTbl.objects.filter(time__date=graphDate).order_by('time')

	for x in range(0,24):
		countPerHour=dayData.filter(time__hour=x).count()
		data.append(countPerHour)

	context={"data":data,"date":graphDate,"prev":prev,"next":nextt,"time":timed}
	return render(request,"graph.php",context)
def roundsGraph(request):
	days=request.GET.get("days","0")
	empid=request.GET.get("empid","")
	floor=request.GET.get("floor","")
	print(floor)
	try:
		days=int(days)
	except Exception as e:
		days=0
	# print(days)

	prev=days-1
	nextt=days+1

	data=[]
	hourlyList=[]
	empDisplayName=''
	# timed=[x for x in range(0,24)]

	graphDate=timezone.localdate() + timedelta(days=days)

	# dayData=pathTbl.objects.filter(time__date=graphDate).order_by('time')
	if floor!='':
		rounds=pathTbl.objects.filter(time__date=graphDate,uuid__placeid__floor=floor).values('empid').annotate(rounds=Count('empid__empid'))
	else:
		rounds=pathTbl.objects.filter(time__date=graphDate).values('empid').annotate(rounds=Count('empid__empid'))
	print(rounds)
	if empid!='':
		if floor!='':
			hourlydata=pathTbl.objects.filter(time__date=graphDate,empid=empid,uuid__placeid__floor=floor)
			empDisplayName=hourlydata.values_list('empid__name',flat=True).distinct()
			for i in range(0,23):
				hourlyList.append(hourlydata.filter(time__hour=i).count())
		else:
			hourlydata=pathTbl.objects.filter(time__date=graphDate,empid=empid)
			empDisplayName=hourlydata.values_list('empid__name',flat=True).distinct()
			for i in range(0,23):
				hourlyList.append(hourlydata.filter(time__hour=i).count())


	# print(rounds)
	# print(rounds.values_list('rounds','empid__name'),'hi')
	names=rounds.values_list('empid__name',flat=True)
	empids=rounds.values_list('empid__empid',flat=True)
	data=rounds.values_list('rounds',flat=True,named=False)
	print(empDisplayName)
	context={"data":list(data),"date":graphDate,"prev":prev,"next":nextt,'names':list(names),'empids':list(empids),'hourly':hourlyList,'days':days,'empName':list(empDisplayName)}


	return render(request,"graph1.php",context)

def saveEmp(request):
	empid=request.GET.get("empid","")
	Ename=request.GET.get("Ename","")
	_phone = request.GET.get("phone","")
	if empid!="" and Ename!="" and _phone!="":
		try:
			emp=empTbl.objects.create(empid=empid,name=Ename,phone=_phone)
			success={"abc":"Success"}
		except:
			success={"abc":'Failed'}
	else:
		success={"abc":'Invalid Entry'}
	return render(request,"add_iss.php",success)
def daterange(start_date, end_date):
    delta = timezone.timedelta(hours=1)
    while start_date < end_date:
        yield start_date#,timezone.localtime(start_date,timezone=timezone.utc)
        start_date += delta
def test(request):
	report_time=timezone.localtime()
	dateTime=request.GET.get("dateTime","")
	if dateTime!='':
		# d,m,y,h=dateTime.split('_')
		tz = timezone.get_current_timezone()
		report_time=tz.localize(datetime.datetime.strptime(dateTime, '%d_%m_%Y_%H'))

	end_time=report_time-timezone.timedelta(minutes=report_time.minute)
	start_time=end_time-timezone.timedelta(hours=12)
	# print(start_time.date(),start_time.time(),end_time.date(),end_time.time(),start_time.tzname())
	filtered_date=pathTbl.objects.filter(time__range=(start_time,end_time)).order_by('time').values_list('time',flat=True)
	# print(filtered_date)
	yes_no=[]
	hrs=[]
	for hr in daterange(start_time,end_time):
		hrs.append(hr)
		countD=filtered_date.filter(time__hour=hr.hour).count()
		# print(countD)
		if countD==0:
			# print(hr,countD)
			yes_no.append('NO')
		else:
			# print(hr,countD,'gotit')
			yes_no.append('YES')
	module_dir = os.path.dirname(__file__)   #get current directory
	file_path = os.path.join(module_dir, 'BeaTrack BLE - Alert Generation Format.xlsx')   #full path to text.

    # data_file = open(file_path , 'rb')   #I used'rb' here since I got an 'gbk' decode error
    # data = data_file.read()
	file_path=os.path.join(module_dir, 'Roster_Report.txt')
	fp = open(file_path,"w")
	# wb = load_workbook(file_path)
	# ws = wb.active
	fp.write('Mr. Prasanta Roy Chowdhury \n')
	fp.write('Site Name: Global Skoda Workshop (Anandapur) \n')
	fp.write('\n')
	fp.write(f'Report Date:	{report_time.strftime("%-d-%b-%y")} \n')
	# fp.write('
	# fp.write('
	# ws['B5']=start_time.date()
	# print(report_time.strftime("%H:%M"))
	if start_time.strftime('%p')=='AM':
		# ws['B6']='A'
		shift_time='A'
	else:
		# ws['B6']='B'
		shift_time='B'

	fp.write(f'Shift for which the report is generated:	{shift_time} \n')
	fp.write(f'Reporting period:	{start_time.strftime("%-I %p (%d-%m-%y)")} to {end_time.strftime("%-I %p (%d-%m-%Y)")} \n')
	fp.write('\n')
	fp.write('Hours                 Patrol Done (Yes/No) \n')
	# ws['B7']=f'{start_time.strftime("%-I %p [%d/%m/%Y]")} to {end_time.strftime("%-I %p [%d/%m/%Y]")}'
	# z=11
	for i, (hrs1, yes_no1) in enumerate(zip(hrs, yes_no)):
		fp.write(f'{hrs1.strftime("%H:%M")}                 {(hrs1+timedelta(hours=1)).strftime("%H:%M")} \n')
		# ws[f'A{z}']=f'{hrs1.strftime("%H:%M")} to {(hrs1+timedelta(hours=1)).strftime("%H:%M")}'
		# ws[f'B{z}']=yes_no1
		# z=z+1
	fp.write('\n')
	fp.write(f'Total number of hours for which patrolling done: {yes_no.count("YES")} \n')
	fp.write(f'Total number of hours for which patrolling not done: {yes_no.count("NO")} \n')
	fp.close()
	# ws['A24']=f'Total number of hours for which patrolling done: {yes_no.count("YES")}'
	# ws['A25']=f'Total number of hours for which patrolling not done: {yes_no.count("NO")}'
	# save_path=os.path.join(module_dir,'Roster_Report.xlsx')

	# print(ws['B5'].value)
	# wb.save(save_path)
	# time_list=range(start_time,12)
	# time_list2=range(1)
	# for x in range(0,24):
	# 	countPerHour=dayData.filter(time__hour=x).count()
	# 	data.append(countPerHour)

	success={"abc":'success'}
	return render(request,"add_iss.php",success)
def test1(request):
	report_time=timezone.localtime()
	empid=request.GET.get("empid","")
	# if dateTime!='':
	# 	# d,m,y,h=dateTime.split('_')
	# 	tz = timezone.get_current_timezone()
	# 	report_time=tz.localize(datetime.datetime.strptime(dateTime, '%d_%m_%Y_%H'))

	end_time=report_time-timezone.timedelta(minutes=report_time.minute)
	start_time=end_time-timezone.timedelta(hours=1)
	# print(start_time.date(),start_time.time(),end_time.date(),end_time.time(),start_time.tzname())
	filtered_date=pathTbl.objects.filter(time__range=(start_time,end_time),empid__empid=empid)
	# print(filtered_date)
	yes_no=''
	print(filtered_date)
	if len(filtered_date)>0:
		yes_no='YES'
	else:
		yes_no='NO'



	success={"abc":yes_no}
	return render(request,"add_iss.php",success)

def storeDataUuid(request):
	# token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	# ttk=request.GET.get("token","")
	uuid=request.GET.get("uuid","")
	site=request.GET.get("site","")
	place=request.GET.get("place","")
	company=request.GET.get("company","")
	other=request.GET.get("other","")
	floor=request.GET.get("floor","")
	building=request.GET.get("building","")
	name=request.GET.get("name","")
	# rssi=request.GET.get("rssi","")
	# time=request.GET.get("time","")
	# print(time)

	# time=datetime.datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
	# print(time)
	
	# empid=request.GET.get("empid","")
	# if ttk!=token:
	# 	success={"abc":"False"}
	# 	return render(request,"add_iss.php",success)
	# else:
	# 	pass
	# print(uuid)
	


	# try:
	if uuidTbl.objects.filter(uuid=uuid):
		success={"abc":'already_exist'}
		return render(request,"add_iss.php",success)
	if uuid!="":		
		placeid=placeTbl.objects.get(siteid__name__icontains=site,floor__icontains=floor,siteid__companyid__name__icontains=company,building__icontains=building)
		print(placeid)
		abc= uuidTbl(uuid=uuid,placeid=placeid,other=other,name=name).save()
		success={"abc":'success'}
	else:
		success={"abc":'Failed'}
	# except:
	# 	success={"abc":'already_exist'}
	return render(request,"add_iss.php",success)
		# abc = pathTbl.
def selectFields(request):
	uuid=request.GET.get("uuid","")
	site=request.GET.get("site","")
	floor=request.GET.get("floor","")
	company=request.GET.get("company","")
	building=request.GET.get("building","")
	name=request.GET.get("name","")

	if company!='' and site=='' and building=='' and floor=='':
		sites_name=siteTbl.objects.filter(companyid__name__icontains=company).values('name')
		data = json.dumps(list(sites_name), cls=DjangoJSONEncoder)
		return HttpResponse(data, content_type="application/json")
	elif company!='' and site!='' and building=='' and floor=='':
		# print(site)
		buildings_name=placeTbl.objects.filter(siteid__companyid__name__icontains=company,siteid__name__icontains=site).values('building')
		# print(buildings_name)
		data = json.dumps(list(buildings_name), cls=DjangoJSONEncoder)
		return HttpResponse(data, content_type="application/json")
	elif company!='' and site!='' and building!='' and floor=='':
		print(site)
		floors_name=placeTbl.objects.filter(siteid__companyid__name__icontains=company,siteid__name__icontains=site,building__icontains=building).values('floor')
		# print(floors_name)
		data = json.dumps(list(floors_name), cls=DjangoJSONEncoder)
		return HttpResponse(data, content_type="application/json")
	# elif company!='' and site!='' and building!='' and floor!='':
	# 	print(site)
	# 	floors_name=placeTbl.objects.filter(siteid__companyid__name__icontains=company,siteid__name__icontains=site,building__icontains=building,floor__icontains=floor).values('floor')
	# 	print(floors_name)
	# 	data = json.dumps(list(floors_name), cls=DjangoJSONEncoder)
	# 	return HttpResponse(data, content_type="application/json")
	else:
		context={"data":'None'}
		return render(request,"lll.php",context)

def storeDataComp(request):
	# token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	# ttk=request.GET.get("token","")
	# uuid=request.GET.get("uuid","")
	# site=request.GET.get("site","")
	# place=request.GET.get("place","")
	company=request.GET.get("company","")
	details=request.GET.get("details","")
	name=request.GET.get("name","")
	# rssi=request.GET.get("rssi","")
	# time=request.GET.get("time","")
	# print(time)

	# time=datetime.datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
	# print(time)
	
	# empid=request.GET.get("empid","")
	# if ttk!=token:
	# 	success={"abc":"False"}
	# 	return render(request,"add_iss.php",success)
	# else:
	# 	pass
	# print(uuid)
	


	try:
		# placeid=companyTbl.objects.get(siteid__name__icontains=site,name__icontains=place,siteid__companyid__name__icontains=company)
		# print(placeid)
		if company!='':
			abc= companyTbl(companyid=company,detail=details,name=company).save()
			success={"abc":'Success'}
		else:
			success={"abc":'Failed blank fields passed'}
	except:
		success={"abc":'Failed'}
	return render(request,"add_iss.php",success)

def storeDataPlace(request):
	# token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	# ttk=request.GET.get("token","")
	# uuid=request.GET.get("uuid","")
	site=request.GET.get("site","")
	# place=request.GET.get("place","")
	# company=request.GET.get("company","")
	about=request.GET.get("about","")
	name=request.GET.get("name","")
	floor=request.GET.get("floor","")
	building=request.GET.get("building","")

	# print(time)

	# time=datetime.datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
	# print(time)
	
	# empid=request.GET.get("empid","")
	# if ttk!=token:
	# 	success={"abc":"False"}
	# 	return render(request,"add_iss.php",success)
	# else:
	# 	pass
	# print(uuid)
	


	try:

		if name!='' and site!='' and building!='' and floor!='':
			siteid=siteTbl.objects.get(name__icontains=site)
			# print(placeid)
			abc= placeTbl(placeid=name,about=about,name=name,siteid=siteid,building=building).save()
			success={"abc":'Success'}
		else:
			success={"abc":'Failed blank fields passed'}
	except:
		success={"abc":'Failed'}
	return render(request,"add_iss.php",success)

def storeDataSite(request):
	# token="1ea94ffa-ef55-4142-b7d8-ce5e45240efe"
	# ttk=request.GET.get("token","")
	loc=request.GET.get("loc","")
	site=request.GET.get("site","")
	# place=request.GET.get("place","")
	company=request.GET.get("company","")
	about=request.GET.get("about","")
	# name=request.GET.get("name","")
	# rssi=request.GET.get("rssi","")
	# time=request.GET.get("time","")
	# print(time)

	# time=datetime.datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
	# print(time)
	
	# empid=request.GET.get("empid","")
	# if ttk!=token:
	# 	success={"abc":"False"}
	# 	return render(request,"add_iss.php",success)
	# else:
	# 	pass
	# print(uuid)
	


	try:
		if company!='' and site!='' and loc!='' :
			companyid=companyTbl.objects.get(name__icontains=company)
			# print(placeid)
			abc= siteTbl(companyid=companyid,about=about,name=site,siteid=site,location=loc).save()
			success={"abc":'Success'}
		else:
			success={"abc":'Failed blank fields passed'}
	except:
		success={"abc":'Failed'}
	return render(request,"add_iss.php",success)



def reportGen(request):
	# report_time=timezone.localtime()
	dateTime=request.GET.get("dateTime","")
	hours=request.GET.get("hours","")
	company=request.GET.get("company","")
	site=request.GET.get("site","")
	place=request.GET.get("place","")
	ivl=request.GET.get("ivl","")
	rnd=request.GET.get("rnd","")
	sft=request.GET.get("sft","")
	place_1=request.GET.get("loc","")

	# target=request.GET.get("target","")
	# target=int(target)
	if hours=='':
		hours=8
	else:
		hours=int(hours)
	if ivl=='':
		ivl=2
	else:
		ivl=int(ivl)
	if rnd=='':
		rnd=2
	else:
		rnd=int(rnd)
	if dateTime!='':
		# d,m,y,h=dateTime.split('_')
		tz = timezone.get_current_timezone()
		report_time=tz.localize(datetime.datetime.strptime(dateTime, '%d_%m_%Y_%H'))
		print(report_time)
	else:
		success={"abc":'Failed blank fields passed'}
		return render(request,"add_iss.php",success)
	# if int(report_time.strftime('%H')) in range(6,13):
	# 	shift='C'
	# elif int(report_time.strftime('%H')) in range(14,21):
	# 	shift='A'
	# elif int(report_time.strftime('%H')) in range(20,7):
	# 	shift='B'
	# else:
	# 	pass


	module_dir = os.path.dirname(__file__)   #get current directory
	file_path1 = os.path.join(module_dir,'yo1.xlsx')

	wb = load_workbook(file_path1)
	ws = wb.active
	# beaconpos=4

	ws.cell(row=2,column=1,value=site)
	ws.cell(row=3,column=1, value=f"Location: {place_1}")



	# print('All empids in place: ',empid_names)
	all_beacons_names=list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site,placeid__siteid__companyid__name__icontains=company,placeid__name__icontains=place).order_by('placeid__about').values_list('name','placeid__floor','placeid__name',named=False))
	# print('All beacons in place: ',all_beacons_names)
	all_beacons_names=list(enumerate(all_beacons_names, start=6))
	total_beacon=len(all_beacons_names)
	if total_beacon<24:
		ws.delete_cols(6+total_beacon,24-total_beacon)
	#print(all_beacons_names)

	for ids3,beacons_name in all_beacons_names:
		# print("program is here")
		ws.cell(row=7, column=ids3,value = 'Beacon \n'+beacons_name[0])
		ws.cell(row=6, column=ids3,value = f'{beacons_name[1]} \n ({beacons_name[2]})')

	# end_time=report_time-timezone.timedelta(minutes=report_time.minute)
	# start_time=end_time-timezone.timedelta(hours=hours)
	end_time=report_time-timezone.timedelta(minutes=report_time.minute)
	start_time=end_time-timezone.timedelta(hours=hours)
	ws.cell(row=5, column=1,value = f"Date : {start_time.strftime('%d %B %Y')} , Shift : {sft}, Timing : {start_time.strftime('%H:%M')} to {end_time.strftime('%H:%M')}")
	start_time1=start_time
	end_time1=end_time
	idx4=6
	for intvrl in range(1,ivl+1):
		end_time1=start_time1+timezone.timedelta(hours=hours//ivl)
	# 	# print(start_time.date(),start_time.time(),end_time.date(),end_time.time(),start_time.tzname())
		filtered_date=pathTbl.objects.filter(time__range=(start_time1,end_time1),uuid__placeid__siteid__name__icontains=site,uuid__placeid__siteid__companyid__name__icontains=company)#.order_by('time')#.values_list('time',flat=True)
		# print(filtered_date)
		# place_filter=placeTbl.objects.filter(siteid__name__icontains=site,siteid__companyid__name__icontains=company,name__icontains=place)
		# print(uuidTbl.objects.filter(placeid=place_filter[0]))
		# floors_names=list(place_filter.values_list('floor',flat=True,named=False))#.distinct('floor')
		# print('All floors in place: ',floors_names)
		empid_names=list(set(filtered_date.values_list('empid',flat=True,named=False)))
		for idx1,empid_name in enumerate(empid_names, start=idx4+2):
			count=0
			# print(f'For emp: {empid_name}')
			# for floors_name in floors_names:
			# 	print(f'	For floor: {floors_name}')
			# 	# print(beacons_names)
			# 	# beacons_floor_names=list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site,placeid__siteid__companyid__name__icontains=company,placeid__floor__icontains=floors_name).values_list('name',flat=True,named=False))

			for idx2,beacons_name in all_beacons_names:
				# ws.cell(row=7, column=beaconpos,value = beacons_name)
				# beaconpos=beaconpos+1
				# if beacons_name not in list(filtered_date.filter(placeid=place_filter).values_list('uuid__name',flat=True,named=False)):
				# 	con
					
				rounds=filtered_date.filter(empid=empTbl.objects.get(empid=empid_name),uuid__name=beacons_name[0]).count()#values('empid').annotate(rounds=Count('empid__empid'))
				# print(f'		Beacon',beacons_name,' Rounds',rounds)
				# if rounds==0:
				# 	continue
				if rounds >= rnd:
				# 	counts=1
					count=count+1
				else:
					count=count+0
				# list(beacons_name in [j for i in mylist for j in i])
				for subarray in all_beacons_names:
					if beacons_name[0] in subarray[1]:
						ws.cell(row=idx1, column=idx2,value='Y' if rounds >= rnd else 'N')
						# print('row ',idx1,'  column ',idx2,' round ',rounds)

				# ws.merge_cells(start_row=6, start_column=5, end_row=6, end_column=len(beacons_floor_names))
			strt_tm=start_time1.strftime('%H:%M')
			endtm=end_time1.strftime('%H:%M')
			#print(strt_tm,endtm)
			ws.cell(row=idx1, column=1,value = f"Slot {intvrl} ({start_time1.strftime('%H:%M')}-{end_time1.strftime('%H:%M')})")
			ws.cell(row=idx1, column=2,value = empTbl.objects.get(empid=empid_name).name)
			# tz = timezone.get_current_timezone()
			# report_time=tz.localize(datetime.datetime.strptime(dateTime, '%d_%m_%Y_%H'))
			ws.cell(row=idx1, column=3,value = timezone.localtime(filtered_date.filter(empid=empTbl.objects.get(empid=empid_name)).order_by('time',)[0].time).strftime('%H:%M'))
			ws.cell(row=idx1, column=4,value = timezone.localtime(filtered_date.filter(empid=empTbl.objects.get(empid=empid_name)).order_by('-time',)[0].time).strftime('%H:%M'))
		
			psr=round((count/total_beacon)*100,2)
			ws.cell(row=idx1, column=5,value = str(psr)+' %')
			idx4=idx1
		start_time1=end_time1

	thick_border = openpyxl.styles.borders.Border(bottom=openpyxl.styles.borders.Side(style='thin'))
	thick_border1 = openpyxl.styles.borders.Border(right=openpyxl.styles.borders.Side(style='thin'))
	thick_border2 = openpyxl.styles.borders.Border(right=openpyxl.styles.borders.Side(style='thin'),bottom=openpyxl.styles.borders.Side(style='thin'))
	thick_border3 = openpyxl.styles.borders.Border(right=openpyxl.styles.borders.Side(style='thin'),
		bottom=openpyxl.styles.borders.Side(style='thin'),top=openpyxl.styles.borders.Side(style='thin'),left=openpyxl.styles.borders.Side(style='thin'))
	if idx4==6:
		idx4=idx4+3
	ws.cell(row=idx4+2,column=1,value='Notes:').font=openpyxl.styles.Font(bold=True)
	ws.cell(row=idx4+3,column=1,value=f'1. We understand that one person will patrol the entire premises atleast {ivl} time(s) each shift').font=openpyxl.styles.Font(bold=True)
	ws.cell(row=idx4+4,column=1,value='2. Patrol success rate to be computed by dividing the number of beacons crossed once by one person per slot against the total number of beacons.').font=openpyxl.styles.Font(bold=True)
	get_row_num=idx4+4
	for col1 in range(1,6+total_beacon):
		ws.cell(row=get_row_num,column=col1).border=thick_border
	file_path2 = os.path.join(module_dir,f'{site} - Patrol Monitoring Report.xlsx')
	get_column_num=6+total_beacon-1

	for rw1 in range(1,get_row_num+1):
		ws.cell(row=rw1,column=get_column_num).border=thick_border1

	ws.cell(row=get_row_num,column=get_column_num).border=thick_border2

	ws.cell(row=6,column=get_column_num).border=thick_border3
	ws.cell(row=7,column=get_column_num).border=thick_border3
	wb.save(file_path2)


	# try:
	# 	strt=start_time.strftime('%d %B %Y')

	# 	mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
	# 	df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

	# 	df_1=df.loc[df['siteid']==site]
	# 	mail_list=df_1.iloc[:,1:].values.tolist()
	# 	mail_list=[element for element in mail_list[0] if '@' in str(element)]
	# 	send_to_list=mail_list

	# 	body=f"Dear Sir,\n\nPlease find attached the patrolling report pertaining to {site} with respect to Shift {sft} as on {strt}.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
	# 	#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
	# 	#send_to_list=['arnab@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
	# 	mail = EmailMessage(site+" | Patrolling Report",body, settings.EMAIL_HOST_USER, send_to_list)
	# 	mail.attach_file(file_path2)
	# 	mail.send()
	# except:
	# 	send_to_list=['arnab@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
	# 	body=f"Dear Sir,\n\nPlease find attached the patrolling report pertaining to {site} with respect to Shift {sft} as on {strt}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
	# 	mail = EmailMessage(site+" | Patrolling Report failed",body, settings.EMAIL_HOST_USER, send_to_list)
	# 	mail.send()

	
	success={"abc":'Success'}
	return render(request,"add_iss.php",success)

	# filtered_date.filter(uuid__name=beacons_names[0])
	# rounds=pathTbl.objects.filter(time__date=graphDate,uuid__placeid__floor=floor).values('empid').annotate(rounds=Count('empid__empid'))


# 	fp.write('Mr. Prasanta Roy Chowdhury \n')
# 	fp.write('Site Name: Global Skoda Workshop (Anandapur) \n')
# 	fp.write('\n')
# 	fp.write(f'Report Date:	{report_time.strftime("%-d-%b-%y")} \n')
# 	# fp.write('
# 	# fp.write('
# 	# ws['B5']=start_time.date()
# 	# print(report_time.strftime("%H:%M"))
# 	if start_time.strftime('%p')=='AM':
# 		# ws['B6']='A'
# 		shift_time='A'
# 	else:
# 		# ws['B6']='B'
# 		shift_time='B'

# 	fp.write(f'Shift for which the report is generated:	{shift_time} \n')
# 	fp.write(f'Reporting period:	{start_time.strftime("%-I %p (%d-%m-%y)")} to {end_time.strftime("%-I %p (%d-%m-%Y)")} \n')
# 	fp.write('\n')
# 	fp.write('Hours                 Patrol Done (Yes/No) \n')
# 	# ws['B7']=f'{start_time.strftime("%-I %p [%d/%m/%Y]")} to {end_time.strftime("%-I %p [%d/%m/%Y]")}'
# 	# z=11
# 	for i, (hrs1, yes_no1) in enumerate(zip(hrs, yes_no)):
# 		fp.write(f'{hrs1.strftime("%H:%M")}                 {(hrs1+timedelta(hours=1)).strftime("%H:%M")} \n')
# 		# ws[f'A{z}']=f'{hrs1.strftime("%H:%M")} to {(hrs1+timedelta(hours=1)).strftime("%H:%M")}'
# 		# ws[f'B{z}']=yes_no1
# 		# z=z+1
# 	fp.write('\n')
# 	fp.write(f'Total number of hours for which patrolling done: {yes_no.count("YES")} \n')
# 	fp.write(f'Total number of hours for which patrolling not done: {yes_no.count("NO")} \n')
# 	fp.close()
# 	# ws['A24']=f'Total number of hours for which patrolling done: {yes_no.count("YES")}'
# 	# ws['A25']=f'Total number of hours for which patrolling not done: {yes_no.count("NO")}'
# 	# save_path=os.path.join(module_dir,'Roster_Report.xlsx')

# 	# print(ws['B5'].value)
# 	# wb.save(save_path)
# 	# time_list=range(start_time,12)
# 	# time_list2=range(1)
# 	# for x in range(0,24):
# 	# 	countPerHour=dayData.filter(time__hour=x).count()
# 	# 	data.append(countPerHour)

def autoreport(request):

	input_str_1=request.GET.get("startdate","")
	input_str_2=request.GET.get("enddate","")
	place_1=request.GET.get("loc","")
	site=request.GET.get("site","")


	# event_list=pd.DataFrame(list(uuidTbl.objects.all().values()))
	# print(event_list)
	# print(type(event_list))

	# input_str_1="2021-11-01"
	start_date=datetime.datetime.strptime(input_str_1, "%Y-%m-%d")
	# input_str_2="2021-11-30"
	end_date=datetime.datetime.strptime(input_str_2, "%Y-%m-%d")

	minutes =30
	hours_added = timezone.timedelta(minutes = minutes)
	start_date=start_date+hours_added

	hours=48
	minutes=30
	hours_added = timezone.timedelta(hours=hours,minutes = minutes)
	end_date=end_date+hours_added
	start__date=timezone.make_aware(start_date)
	end_date=timezone.make_aware(end_date)

	df_path=pd.DataFrame(list(pathTbl.objects.all().values()))

	# print(start_date,end_date)

	# df_path=pd.DataFrame(list(pathTbl.objects.filter(time__range=(start_date,end_date)).values()))

	df_path['time']=df_path['time'].dt.tz_localize(None)

	# print(df_path)

	# print(df_path.head())
	# print(df_path['time'].dtypes)

	# test_1=df_path.iloc[1,5:6]
	# print(test_1[5:8])
	# print(df_path['time'].dtypes)



	#df_uuid=pd.DataFrame(list(uuidTbl.objects.all().values()))

	df_uuid=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site).values()))
	print(df_uuid)

	module_dir = os.path.dirname(__file__)
	if site=='Utalika-Residence':
		file_path1 = os.path.join(module_dir,"Utalika_report.xlsx")
	else:
		file_path1 = os.path.join(module_dir,"AMRI Patrol Report -1.xlsx") #get current directory
	
	start_date=request.GET.get("startdate","")
	end_date=request.GET.get("enddate","")
	final_dict=data_generation.data_generation(input_str_1,input_str_2,df_uuid,df_path,12,2,8)


	wb_obj = openpyxl.load_workbook(file_path1)

	ws_obj=wb_obj["Performance Report"]

	dt_1=datetime.datetime.strptime(input_str_1, "%Y-%m-%d").strftime("%d %B %Y")
	dt_2=datetime.datetime.strptime(input_str_2, "%Y-%m-%d").strftime("%d %B %Y")


	headline="Sumarry Report for "+dt_1+" to "+dt_2
	ws_obj.cell(row=2,column=1).value=site
	ws_obj.cell(row=4,column=1).value=headline
	
	#print(final_dict)

	if site=='Utalika-Residence':
		day_summary_dict=final_dict['night_wise_summary']
		value_list_day=list(day_summary_dict.values())
		row_day=9
		column_day=3
		k=0
		while k<len(value_list_day):
			ws_obj.cell(row=row_day, column=column_day).value=value_list_day[k]
			row_day=row_day+1
			k=k+1

		beacon_dict=final_dict['beacon summary']
		marklist = sorted(beacon_dict.items(), key=lambda x:x[1])
		sortdict = dict(marklist)
		# print(sortdict)

		res_beacon = dict(reversed(list(sortdict.items())))

		key_list_beacon=list(res_beacon.keys())
		value_list_beacon=list(res_beacon.values())
		perc_list=final_dict['beacon_perc']
		perc_list.sort(reverse=True)
		# print(perc_list)

		ws_obj.cell(row=20,column=1).value=f"Part II: Frequency of location coverage within {site} premises"

		k=0
		row_beacon=23
		col_beacon=2
		col_val=3
		col_perc=4
		while k < len(res_beacon):
			ws_obj.cell(row=row_beacon,column=col_beacon).value=key_list_beacon[k]
			ws_obj.cell(row=row_beacon,column=col_val).value=value_list_beacon[k]
			ws_obj.cell(row=row_beacon,column=col_perc).value=perc_list[k]
			k=k+1
			row_beacon=row_beacon+1

	else:



		day_summary_dict=final_dict['day_wise_summary']
		night_summary_dict=final_dict['night_wise_summary']
		# print(night_summary_dict)

		# print(day_summary_dict)
		

		value_list_day=list(day_summary_dict.values())
		row_day=9
		column_day=3
		k=0
		while k<len(value_list_day):
			ws_obj.cell(row=row_day, column=column_day).value=value_list_day[k]
			row_day=row_day+1
			k=k+1

		value_list_night=list(night_summary_dict.values())
		k=0
		row_night=23
		column_night=3
		while k<len(value_list_night):
			ws_obj.cell(row=row_night, column=column_night).value=value_list_night[k]
			row_night=row_night+1
			k=k+1

		beacon_dict=final_dict['beacon summary']
		marklist = sorted(beacon_dict.items(), key=lambda x:x[1])
		sortdict = dict(marklist)
		# print(sortdict)

		res_beacon = dict(reversed(list(sortdict.items())))

		key_list_beacon=list(res_beacon.keys())
		value_list_beacon=list(res_beacon.values())
		perc_list=final_dict['beacon_perc']
		perc_list.sort(reverse=True)
		# print(perc_list)

		ws_obj.cell(row=34,column=1).value=f"Part III: Frequency of location coverage within {site} premises"

		k=0
		row_beacon=37
		col_beacon=2
		col_val=3
		col_perc=4
		while k < len(res_beacon):
			ws_obj.cell(row=row_beacon,column=col_beacon).value=key_list_beacon[k]
			ws_obj.cell(row=row_beacon,column=col_val).value=value_list_beacon[k]
			ws_obj.cell(row=row_beacon,column=col_perc).value=perc_list[k]
			k=k+1
			row_beacon=row_beacon+1

	# if site=="Utalika-Residence":
	# 	ws_obj.delete_rows(37+len(df_uuid),37+24)
		# ws_obj.delete_rows(6,18)



	try:
		month_name_obj=datetime.datetime.strptime(input_str_1[5:7], "%m")
		month_name=month_name_obj.strftime("%B")
		filenam=str(site)+" Monthly Patrol Performance Report_"+str(month_name)+"_"+input_str_1[0:4]+".xlsx"
	except:
		filenam=str(site)+" Monthly Patrol Performance Report.xlsx"

	file_path2 = os.path.join(module_dir,filenam)

	wb_obj.save(file_path2)

	try:
		# mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		# df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# # print(df)

		# df_1=df.loc[df['siteid']==site]
		# mail_list=df_1.iloc[:,1:].values.tolist()
		# mail_list=[element for element in mail_list[0] if '@' in str(element)]
		# send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{site}- Monthly Patrol Performance Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {site} for the period of {dt_1} to {dt_2}.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the monthly patrol performance report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(file_path2)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {site} for the period of {dt_1} to {dt_2}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		mail = EmailMessage(site+" | Monthly Report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	# print(len(event_list))


	success={"abc":'Success'}
	return render(request,"add_iss.php",success)


def qrDashboard(request):
	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	site_code = request.GET.get('sitecode',"")
	seckey=request.GET.get("seckey","")
	viewmode=request.GET.get("view","")
	starthr=request.GET.get("starthr","")
	placeselect=request.GET.get("placeselect","")
	gap=request.GET.get("gap","")

	if starthr=="":
		starthr=8
	else:
		try:
			starthr=int(starthr)
		except:
			starthr=8

	date_filter1=request.GET.get("date_inp","")

	if date_filter1=="":

		time1=datetime.now()
		# time1=time1+timedelta(hours=5,minutes=30)
		time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		endhr=24-starthr
		time2_11=time2_1-timedelta(hours=endhr)
		time2_1=time2_1+timedelta(hours=starthr)
		tz = timezone.get_current_timezone()
		end_time1=tz.localize(time2_1)
		start_time1=tz.localize(time2_11)
	else:
		tz = timezone.get_current_timezone()

		start_time1=tz.localize(datetime.strptime(date_filter1, '%Y-%m-%d'))
		if site_code=="Teloijan-Tea":
			install_time=tz.localize(datetime.strptime("2022-09-16", '%Y-%m-%d'))
			if start_time1<install_time:
				return HttpResponse('''<br><br><br><br><br><br><br><br><br><br><br><h1>Installtion of qr codes were not done on the selected date. Please select other date.</h1>''')
		end_time1=start_time1+timedelta(hours=32)
		start_time1=start_time1+timedelta(hours=8)

	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	if site_code=="Teloijan-Tea":
		strt_tm1=start_time1+timedelta(hours=12)
		dt_1=datetime.strftime(strt_tm1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

		
	else:
		dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code).values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})
	# print(all_path)

	

	# df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')


	# print(df_all["time"])
	# df_all['time']=df_all['time']+timedelta(hours =5, minutes=30)
	

	dict_hour={"time":[]}
	# all_beacons_names["full1"]=all_beacons_names["placeid_id"][0:len(all_beacons_names["placeid_id"])-len(site_code)-2]

	# print(all_beacons_names["full1"])

	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]


	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	df_all=pd.merge(all_path,all_beacons_names,on='uuid_id')

	df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

	df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
	df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %Y %I:%M %p')
	df_all["dateonly"]=df_all["time"].dt.date
	df_all["hour"]=df_all["time"].dt.hour
	df_all["minute"]=df_all["time"].dt.minute


	if viewmode=="json":
		end_date1=datetime.now()
		end_date1=time1-timedelta(days=1,hours=time1.hour,minutes=time1.minute)

		end_date1=tz.localize(end_date1)
		st_tm1=end_date1+timedelta(hours=starthr)
		end_tm1=end_date1+timedelta(days=1,hours=starthr)

		# date_str=datetime.strftime(end_date1,"%d-%b")
		# date_list.append(date_str)

		all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))




	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(st_tm1,end_tm1)).values()))

		if len(all_path)==0:
			succ_list.append(0)

		else:

			df_all=pd.merge(all_path,all_beacons_names,on='uuid_id')
			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %Y %I:%M %p')
			df_all["dateonly"]=df_all["time"].dt.date
			df_all["hour"]=df_all["time"].dt.hour
			df_all["minute"]=df_all["time"].dt.minute

			suc_rate=get_success_rate(df_all,24,name_list,starthr,gap,site_code)

			suc_rate=round(suc_rate,2)


		return JsonResponse({"foo":suc_rate})


	# print(len(nameonly_list))

	maxval=0
	minval=100
	maxpl=""
	minpl=""

	df_maxmin=df_all.groupby("fullname")

	for n1,grps in df_maxmin:
		if len(grps)>maxval:
			maxval=len(grps)
			maxpl=n1

		if len(grps)<=minval:
			minval=len(grps)
			minpl=n1


	# for ele11 in name_list:
	# 	df_max=

	for ele in name_list:
		dict_hour[ele]=[]

	# print(df_all)
	# print(dict_hour)

	start_hr=starthr
	end_hr=starthr+1


	i=0
	total_slot=24
	counter=0

	uuid_list=all_beacons_names["uuid_id"].tolist()

	while i<total_slot:


		if start_hr==23:
			df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
		else:
			df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]
		# print(len(df_all_1))


		if start_hr<10:
			strt="0"+str(start_hr)
		else:
			strt=str(start_hr)

		if end_hr<10:
			ent="0"+str(end_hr)
		else:
			ent=str(end_hr)



		if gap=="1hr":
			time_stamp1=strt+":00-"+ent+":00 Hours"
			dict_hour["time"].append(time_stamp1)

		else:

			time_stamp1=strt+":00-"+strt+":30 Hours"
			dict_hour["time"].append(time_stamp1)
			time_stamp2=strt+":30-"+ent+":00 Hours"
			dict_hour["time"].append(time_stamp2)

		if len(df_all_1)==0:
			
			for idx,ele1 in enumerate(name_list):
				# counter=counter+1
				
				# print("program is here")
				if gap=="1hr":
					dict_hour[ele1].append("color:#ff0000")
					# dict_hour[ele1].append("color:#ff0000")
				else:
					dict_hour[ele1].append("color:#ff0000")
					dict_hour[ele1].append("color:#ff0000")




		else:
			


			for idx1,ele2 in enumerate(name_list):


				if gap=="1hr":
					df_all_2=df_all_1.loc[df_all_1["fullname"]==ele2]

					if len(df_all_2)==0:
						dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
					else:
						dict_hour[ele2].append("color:#00ff00;")


				# counter=counter+1
				else:
				
					df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



					if len(df_all_2)==0:
						dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
					else:
						dict_hour[ele2].append("color:#00ff00;")
						# counter=counter+1


					df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

					if len(df_all_3)==0:
						dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
					else:
						dict_hour[ele2].append("color:#00ff00;")
					# counter=counter+1


		if start_hr==23:
			start_hr=0
		else:
			start_hr=start_hr+1

		if end_hr==23:
			end_hr=0

		else:
			end_hr=end_hr+1

		i=i+1


	df_final=pd.DataFrame(dict_hour)
	# print(counter)
	# print(len(dict_hour))

	# for key in dict_hour:
	# 	print(len(dict_hour[key]),key)

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)


	records=df_final.to_records(index=False)
	result=tuple(records)

	if site_code=="Teloijan-Tea":
		result=result[12:24]

	# print(result)

	name_list1=[]


	for ele10 in name_list:
		x1=ele10[0:(len(ele10)-len(site_code)-2)]
		name_list1.append(x1)




	#succress rate calculation

	time1=datetime.now()

	end_date1=time1-timedelta(days=1,hours=time1.hour,minutes=time1.minute)

	end_date1=tz.localize(end_date1)

	date_list=[]
	succ_list=[]

	j=0
	while j<7:
		st_tm1=end_date1+timedelta(hours=starthr)
		end_tm1=end_date1+timedelta(days=1,hours=starthr)

		date_str=datetime.strftime(end_date1,"%d-%b")
		date_list.append(date_str)

		all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))




	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(st_tm1,end_tm1)).values()))

		if len(all_path)==0:
			succ_list.append(0)

		else:

			df_all=pd.merge(all_path,all_beacons_names,on='uuid_id')
			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %Y %I:%M %p')
			df_all["dateonly"]=df_all["time"].dt.date
			df_all["hour"]=df_all["time"].dt.hour
			df_all["minute"]=df_all["time"].dt.minute

			suc_rate=get_success_rate(df_all,24,name_list,starthr,gap,site_code)

			suc_rate=round(suc_rate,2)
			succ_list.append(suc_rate)

		end_date1=end_date1-timedelta(days=1)

		j=j+1


	# print(date_list,succ_list)

		# all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})
















	# print(df_all["time"])



	# print("Sitecode is--->>>", siteid)
	# start = datetime(2022, 7, 26, 00, 00, 00)
	# end = datetime(2022, 7, 27, 00, 00, 00)
	# siteid = "NIS-Head Office"
	maxpl=maxpl[0:(len(maxpl)-len(site_code)-2)]
	minpl=minpl[0:(len(minpl)-len(site_code)-2)]

	date_list=date_list[::-1]
	succ_list=succ_list[::-1]

	# print(date_list,succ_list)

	if site_code=="Teloijan-Tea":
		onlynight=True
	else:
		onlynight=False
	# print(onlynight)

	context={"namelist":name_list1,"dt_1":dt_1,"x":data3,"data":result,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"date_list":date_list,"succ_list":succ_list,
	"filter_date":date_filter1,"starthr":starthr,"maxpl":maxpl,"minpl":minpl,"gap":gap,"onlynight":onlynight}
	
	
	return render(request, "qrDashboardMobile.html",context)


def timestamp(request):
	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	site_code = request.GET.get('sitecode',"")
	seckey=request.GET.get("seckey","")
	viewmode=request.GET.get("view","")
	starthr=request.GET.get("starthr","")
	placeselect=request.GET.get("placeselect","")
	gap=request.GET.get("gap","")

	if starthr=="":
		starthr=8
	else:
		try:
			starthr=int(starthr)
		except:
			starthr=8

	date_filter1=request.GET.get("date_inp","")

	if date_filter1=="":

		time1=datetime.now()
		# time1=time1+timedelta(hours=5,minutes=30)
		time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		endhr=24-starthr
		time2_11=time2_1-timedelta(hours=endhr)
		time2_1=time2_1+timedelta(hours=starthr)
		tz = timezone.get_current_timezone()
		end_time1=tz.localize(time2_1)
		start_time1=tz.localize(time2_11)
	else:
		tz = timezone.get_current_timezone()

		start_time1=tz.localize(datetime.strptime(date_filter1, '%Y-%m-%d'))

		if site_code=="Teloijan-Tea":
			install_time=tz.localize(datetime.strptime("2022-09-16", '%Y-%m-%d'))
			if start_time1<install_time:
				return HttpResponse('''<br><br><br><br><br><br><br><br><br><br><br><h1>Installtion of qr codes were not done on the selected date. Please select other date.</h1>''')
		end_time1=start_time1+timedelta(hours=32)
		start_time1=start_time1+timedelta(hours=8)

	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	if site_code=="Teloijan-Tea":
		strt_tm1=start_time1+timedelta(hours=12)
		dt_1=datetime.strftime(strt_tm1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

		
	else:
		dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code).values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})
	# print(all_path)

	

	# df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')


	# print(df_all["time"])
	# df_all['time']=df_all['time']+timedelta(hours =5, minutes=30)
	

	dict_hour={"time":[]}
	# all_beacons_names["full1"]=all_beacons_names["placeid_id"][0:len(all_beacons_names["placeid_id"])-len(site_code)-2]

	# print(all_beacons_names["full1"])

	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]


	# if placeselect=="":

	name_list=all_beacons_names["fullname"].tolist()

	nameonly_list=all_beacons_names["name"].tolist()

	# else:

		# name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		# nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

	all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
	all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

	df_all=pd.merge(df_all1,all_emp,on='empid_id')


	df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

	df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
	df_all["datetime1"]=df_all["time"].dt.strftime('%d-%b %I:%M %p')
	df_all["dateonly"]=df_all["time"].dt.date
	df_all["hour"]=df_all["time"].dt.hour
	df_all["minute"]=df_all["time"].dt.minute

	df_all=df_all.sort_values(by=['time'],ascending=True)

	if placeselect=="":


		json_records3 = df_all.reset_index().to_json(orient ='records')
		data3 = []
		data3 = json.loads(json_records3)

	else:
		df_all=df_all.loc[df_all["name"]==placeselect]
		json_records3 = df_all.reset_index().to_json(orient ='records')
		data3 = []
		data3 = json.loads(json_records3)

	context={"dt_1":dt_1,"gap":gap,"starthr":starthr,"vs":data3,"sites":site_code,"placelist":nameonly_list,"filter_date":date_filter1}
	
	
	return render(request, "timestamp.html",context)






def get_success_rate(df_all,total_slot,name_list,starthr,gap,site_code):

	start_hr=starthr
	end_hr=starthr+1

	total_count=len(name_list)*total_slot

	counter=0
	i=0

	while i<total_slot:


		if start_hr==23:
			df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
		else:
			# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
			df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]
		# print(len(df_all_1))


		# if start_hr<10:

		if len(df_all_1)==0:
			pass



		else:

			

			


			for idx1,ele2 in enumerate(name_list):
				# counter=counter+1

				if gap=="1hr":
					df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



					if len(df_all_2)==0:
						# dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
						pass
					else:
						# dict_hour[ele2].append("color:#00ff00;")
						counter=counter+1

				else:
					df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



					if len(df_all_2)==0:
						# dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
						pass
					else:
						# dict_hour[ele2].append("color:#00ff00;")
						counter=counter+1


					df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

					if len(df_all_3)==0:
						pass
						# dict_hour[ele2].append("color:#ff0000")
						# counter=counter+1
					else:
						counter=counter+1

		if start_hr==23:
			start_hr=0
		else:
			start_hr=start_hr+1

		if end_hr==23:
			end_hr=0

		else:
			end_hr=end_hr+1

		i=i+1



	if site_code=="Teloijan-Tea":
		success_rate=(counter/total_count)*100*2
	else:
		success_rate=(counter/total_count)*100

	return success_rate



def successrate(request):
	site_code = request.GET.get('sitecode',"")
	seckey=request.GET.get("seckey","")
	viewmode=request.GET.get("view","")
	starthr=request.GET.get("starthr","")
	placeselect=request.GET.get("placeselect","")
	gap=request.GET.get("gap","")


	if starthr=="":
		starthr=8
	else:
		try:
			starthr=int(starthr)
		except:
			starthr=8


	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})
	# print(len(all_beacons_names))
	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]


	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()

	tz=timezone.get_current_timezone()


	time1=datetime.now()

	end_date1=time1-timedelta(days=1,hours=time1.hour,minutes=time1.minute)

	end_date1=tz.localize(end_date1)

	date_list=[]
	succ_list=[]


	j=0
	while j<30:
		st_tm1=end_date1+timedelta(hours=starthr)
		end_tm1=end_date1+timedelta(days=1,hours=starthr)

		date_str=datetime.strftime(end_date1,"%d-%B")
		date_list.append(date_str)

		all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))




	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(st_tm1,end_tm1)).values()))

		if len(all_path)==0:
			succ_list.append(0)

		else:

			df_all=pd.merge(all_path,all_beacons_names,on='uuid_id')
			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %Y %I:%M %p')
			df_all["dateonly"]=df_all["time"].dt.date
			df_all["hour"]=df_all["time"].dt.hour
			df_all["minute"]=df_all["time"].dt.minute

			suc_rate=get_success_rate(df_all,24,name_list,starthr,gap,site_code)

			suc_rate=round(suc_rate,2)
			succ_list.append(suc_rate)

		end_date1=end_date1-timedelta(days=1)

		j=j+1

	if site_code=="Teloijan-Tea":
		onlynight=True
	else:
		onlynight=False

	context={"succ_list":succ_list,"date_list":date_list,"sites":site_code,"starthr":starthr,"gap":gap,"onlynight":onlynight}




	return render(request, "allsuccess.html",context)









def livedata(request):

	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	site_code = request.GET.get('sitecode',"")
	seckey=request.GET.get("seckey","")
	viewmode=request.GET.get("view","")
	starthr=request.GET.get("starthr","")
	placeselect=request.GET.get("placeselect","")
	gap=request.GET.get("gap","")

	if starthr=="":
		starthr=8
	else:
		try:
			starthr=int(starthr)
		except:
			starthr=8

	# date_filter1=request.GET.get("date_inp","")

	# if date_filter1=="":

	tz = timezone.get_current_timezone()

	time1=datetime.now()
	# time1=tz.localize(time1
	# time1=time1+timedelta(hours=5,minutes=30)
	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_11=time2_1+timedelta(hours=starthr)

	endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time1)
	start_time1=tz.localize(time2_11)

	if int(end_time1.hour)<20 and site_code=="Teloijan-Tea" and int(end_time1.hour)>=8:
		msg="The post monitoring on this site is done during night shift. Please check live data after 08:00 p.m"
		context={"sites":site_code,"starthr":starthr,"gap":gap,"msg":msg}

		return render(request, "qrlive.html",context)



	# end_time1=timezone.localtime(end_time1)
	# start_time1=timezone.localtime(start_time1)


	# end_time1=time1
	# start_time1=time2_11

	# else:
	# 	tz = timezone.get_current_timezone()

	# 	start_time1=tz.localize(datetime.strptime(date_filter1, '%Y-%m-%d'))
	# 	end_time1=start_time1+timedelta(hours=32)
	# 	start_time1=start_time1+timedelta(hours=8)

	if site_code=="Teloijan-Tea":
		if int(time1.hour)>=20:
			start_time1=start_time1-timedelta(hours=start_time1.hour)+timedelta(hours=20)
			starthr=20
		elif int(time1.hour)<8:
			start_time1=start_time1-timedelta(hours=start_time1.hour)+timedelta(hours=20)
			start_time1=start_time1-timedelta(days=1)
			starthr=20
		else:
			start_time1=start_time1-timedelta(hours=start_time1.hour)+timedelta(hours=20)
			starthr=20


	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code).values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})
	# print(all_path)

	

	# df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')


	# print(df_all["time"])
	# df_all['time']=df_all['time']+timedelta(hours =5, minutes=30)
	

	dict_hour={"time":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]

	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	for ele in name_list:
		dict_hour[ele]=[]

	if len(all_path)==0:
		start_hr=starthr
		end_hr=starthr+1
		i=0
		now_hour=int(time1.hour)

		if now_hour>=20:
			total_slot=abs(now_hour-start_hr+1)
		elif now_hour<8:
			total_slot=abs((24-starthr)+now_hour+1)
		else:
			total_slot=abs(now_hour-start_hr+1)

		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:


			# if start_hr==23:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
			# else:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
			# print(len(df_all_1))


			# if start_hr<10:




			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)

			for idx,ele1 in enumerate(name_list):
				# counter=counter+1
				
				# print("program is here")
				if gap=="1hr":
					dict_hour[ele1].append("color:#ff0000")
				else:

					dict_hour[ele1].append("color:#ff0000")
					dict_hour[ele1].append("color:#ff0000")

			if start_hr==23:
				start_hr=0
			else:
				start_hr=start_hr+1

			if end_hr==23:
				end_hr=0

			else:
				end_hr=end_hr+1

			i=i+1





	else:
		df_all=pd.merge(all_path,all_beacons_names,on='uuid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %Y %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute



		start_hr=starthr
		end_hr=starthr+1


		i=0
		now_hour=int(time1.hour)
		if now_hour>=20:
			total_slot=abs(now_hour-start_hr+1)
		elif now_hour<8:
			total_slot=abs((24-starthr)+now_hour+1)
		else:
			total_slot=abs(now_hour-start_hr+1)
		# total_slot=abs(now_hour-start_hr+1)
		# print("this is total slot", total_slot)
		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:


			if start_hr==23:
				df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
			else:
				# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
				df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]
			# print(len(df_all_1))


			# if start_hr<10:




			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)

			if len(df_all_1)==0:
				
				for idx,ele1 in enumerate(name_list):
					# counter=counter+1
					
					# print("program is here")
					if gap=="1hr":
						dict_hour[ele1].append("color:#ff0000")
					else:

						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")



			else:
				


				for idx1,ele2 in enumerate(name_list):
					# counter=counter+1

					if gap=="1hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							dict_hour[ele2].append("color:#00ff00;")

					else:

					
					
						df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							dict_hour[ele2].append("color:#00ff00;")
							# counter=counter+1


						df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

						if len(df_all_3)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							dict_hour[ele2].append("color:#00ff00;")
						# counter=counter+1


			if start_hr==23:
				start_hr=0
			else:
				start_hr=start_hr+1

			if end_hr==23:
				end_hr=0

			else:
				end_hr=end_hr+1

			i=i+1


	df_final=pd.DataFrame(dict_hour)
	# print(counter)
	# print(len(dict_hour))

	# for key in dict_hour:
	# 	print(len(dict_hour[key]),key)

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)


	records=df_final.to_records(index=False)
	result=tuple(records)
	result=result[::-1]

	name_list1=[]


	for ele10 in name_list:
		x1=ele10[0:(len(ele10)-len(site_code)-2)]
		name_list1.append(x1)

	# print(result)













	# print(df_all["time"])



	# print("Sitecode is--->>>", siteid)
	# start = datetime(2022, 7, 26, 00, 00, 00)
	# end = datetime(2022, 7, 27, 00, 00, 00)
	# siteid = "NIS-Head Office"

	context={"namelist":name_list1,"dt_1":dt_1,"data":result,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap}
	
	
	return render(request, "qrlive.html",context)


def employeeView(request):
	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	# print("-----Emp view called-------------")
	siteid = request.GET.get('sitecode')
	empid = request.GET.get('emp')
	Name = empTbl.objects.get(empid=empid).name
	# print("Emp Name", Name)

	# print("Sitecode is--->>>", siteid, "Empid is-->>", empid)
	# siteid = "NIS-Head Office"
	time_list = [(8, 9), (10, 11), (12, 13), (14, 15), (16, 17), (18, 19), (20, 21), (22, 23), (0, 1), (2, 3), (4, 5), (6, 7)]
	# splitTimeMonitorListEmp = []
	dateNow = date.today()
	# for i in time_list:
	# 	splitTimeStart = datetime(2022, 7, 26, i[0], 00, 00)
	# 	splitTimeEnd = datetime(2022, 7, 26, i[1], 59, 59)
	# 	empWiseData = pathTbl.objects.filter(time__gte = splitTimeStart, time__lte = splitTimeEnd, empid_id=empid)
	# 	splitTimeMonitorListEmp.append(len(empWiseData))

	end_dt = date.today()
	start_dt = end_dt - timedelta(days=29)
	last30 = start_dt.strftime("%Y %m %d") + " to " + end_dt.strftime("%Y %m %d")
	oneMonthListEmp = []
	date_list = []

	# print("------Start Dt--->>>", start_dt)
	# print("------End Dt--->>>", end_dt)
	for dt in daterange(start_dt, end_dt):
		date_list.append(dt.strftime("%Y %m %d"))
	
	# print(date_list)
	for i in time_list:
		date_data = []
		for dt in date_list:
			# print("--------Dt---------", dt)
			dateof = dt.split()
			splitTimeStartDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[0], 00, 00)
			splitTimeEndDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[1], 59, 59)
			empWiseData = pathTbl.objects.filter(time__gte = splitTimeStartDay, time__lte = splitTimeEndDay, empid_id=empid)
			date_data.append(len(empWiseData))
		# print("date_data-->>", date_data)
		oneMonthListEmp.append(sum(date_data))
	

	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	# print("----------------ActiveQR Code-------->>>", len(activeQR)) #activeQR,
	place_list = [i.name for i in activeQR]
	# print("----------Place Name-->>>", place_list)
	uuidList = []
	for place in place_list:	
		qrList = uuidTbl.objects.get(placeid_id=place)
		# print("qrList-->>", qrList.uuid)
		uuidList.append(qrList.uuid)

	empIdList = []
	empNameList = []
	for uid in uuidList:
		dateNow = date.today()
		dateFrom = dateNow - timedelta(days=29)
		empId = pathTbl.objects.filter(uuid=uid, time__range = (dateFrom, dateNow))
		for emp in empId:
			if emp.empid_id not in empIdList:
				empIdList.append(emp.empid_id)
	for emp in empIdList:
		empName = empTbl.objects.get(empid=emp).name
		empNameList.append(empName)
	uuid_name_list = list(zip(uuidList, place_list))
	empid_name_list = list(zip(empIdList, empNameList))

	print("empid_name_list before----->>>", empid_name_list)
	for i in empid_name_list:
		if i[1] in notEmp:
			empid_name_list.remove(i)
	print("empid_name_list After----->>>", empid_name_list)

	totalPostMonitored = sum(oneMonthListEmp)
	# print("-------------------------------OneDayData---------------------------")
	# print(OneDayData)
	# print("----------------------------------END--------------------------------")

	mostof30 = max(oneMonthListEmp)
	leastof30 = min(oneMonthListEmp)
	time_val_list = ['8AM to 10AM','10AM to 12NOON','12NOON to 2PM','2PM to 4PM','4PM to 6PM','6PM to 8PM','8PM to 10PM','10PM to 00Hour','00HOUR to 2AM','2AM to 4AM','4AM to 6AM','6AM to 8AM']
	mostlast30time = time_val_list[oneMonthListEmp.index(mostof30)]
	leastlast30time = time_val_list[oneMonthListEmp.index(leastof30)]

	context = {
		# 'data' : splitTimeMonitorListEmp,
		'data' : oneMonthListEmp,
		'start':start_dt,
		'end': end_dt,
		'totalDays': len(date_list),
		'sitecode' : siteid,
		'last30': last30,
		"name": Name,
		'uuidList' : uuid_name_list,
		'empdata' : empid_name_list,
		'totalDays': len(date_list),
		'total': totalPostMonitored,
		"mostlast30time": mostlast30time,
		"leastlast30time": leastlast30time,
	}

	return render(request, "Post\\empViewMobile.html", context)

def postView(request):
	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	# print("-----Post view called-------------")
	siteid = request.GET.get('sitecode')
	postid = request.GET.get('placeid')
	postName = uuidTbl.objects.get(uuid=postid).placeid_id
	# print("PostName", postName)

	# print("Sitecode is--->>>", siteid, "Place Id is-->>", postid)
	# siteid = "NIS-Head Office"
	time_list = [(8, 9), (10, 11), (12, 13), (14, 15), (16, 17), (18, 19), (20, 21), (22, 23), (0, 1), (2, 3), (4, 5), (6, 7)]
	# splitTimeMonitorListPost = []
	dateNow = date.today()
	# for i in time_list:
	# 	splitTimeStart = datetime(2022, 7, 26, i[0], 00, 00)
	# 	splitTimeEnd = datetime(2022, 7, 26, i[1], 59, 59)
	# 	postWiseData = pathTbl.objects.filter(time__gte = splitTimeStart, time__lte = splitTimeEnd, uuid_id=postid)
	# 	splitTimeMonitorListPost.append(len(postWiseData))


	end_dt = date.today()
	start_dt = end_dt - timedelta(days=29)
	last30 = start_dt.strftime("%Y %m %d") + " to " + end_dt.strftime("%Y %m %d")
	oneMonthListPost = []
	date_list = []

	# print("------Start Dt--->>>", start_dt)
	# print("------End Dt--->>>", end_dt)
	for dt in daterange(start_dt, end_dt):
		date_list.append(dt.strftime("%Y %m %d"))

	print(date_list)
	for i in time_list:
		date_data = []
		for dt in date_list:
			# print("--------Dt---------", dt)
			dateof = dt.split()
			splitTimeStartDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[0], 00, 00)
			splitTimeEndDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[1], 59, 59)
			postWiseData = pathTbl.objects.filter(time__gte = splitTimeStartDay, time__lte = splitTimeEndDay, uuid_id=postid)
			date_data.append(len(postWiseData))
		# print("date_data-->>", date_data)
		oneMonthListPost.append(sum(date_data))	

	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	# print("----------------ActiveQR Code-------->>>", len(activeQR)) #activeQR,
	place_list = [i.name for i in activeQR]
	# print("----------Place Name-->>>", place_list)
	uuidList = []
	for place in place_list:	
		qrList = uuidTbl.objects.get(placeid_id=place)
		# print("qrList-->>", qrList.uuid)
		uuidList.append(qrList.uuid)

	empIdList = []
	empNameList = []
	for uid in uuidList:
		dateNow = date.today()
		dateFrom = dateNow - timedelta(days=29)
		empId = pathTbl.objects.filter(uuid=uid, time__range = (dateFrom, dateNow))
		for emp in empId:
			if emp.empid_id not in empIdList:
				empIdList.append(emp.empid_id)
	for emp in empIdList:
		empName = empTbl.objects.get(empid=emp).name
		empNameList.append(empName)

	uuid_name_list = list(zip(uuidList, place_list))
	empid_name_list = list(zip(empIdList, empNameList))

	print("empid_name_list before----->>>", empid_name_list)
	for i in empid_name_list:
		if i[1] in notEmp:
			empid_name_list.remove(i)
	print("empid_name_list After----->>>", empid_name_list)

	totalPostMonitored = sum(oneMonthListPost)
	# print("-------------------------------OneDayData---------------------------")
	# print(OneDayData)
	# print("----------------------------------END--------------------------------")

	mostof30 = max(oneMonthListPost)
	leastof30 = min(oneMonthListPost)
	time_val_list = ['8AM to 10AM','10AM to 12NOON','12NOON to 2PM','2PM to 4PM','4PM to 6PM','6PM to 8PM','8PM to 10PM','10PM to 00Hour','00HOUR to 2AM','2AM to 4AM','4AM to 6AM','6AM to 8AM']
	mostlast30time = time_val_list[oneMonthListPost.index(mostof30)]
	leastlast30time = time_val_list[oneMonthListPost.index(leastof30)]

	context = {
		# 'data' : splitTimeMonitorListPost,
		'data' : oneMonthListPost,
		'start':start_dt,
		'end': end_dt,
		'last30': last30,
		'totalDays': len(date_list),
		'sitecode' : siteid,
		"postid": postName,
		'uuidList' : uuid_name_list,
		'empdata' : empid_name_list,
		'totalDays': len(date_list),
		'total': totalPostMonitored,
		"mostlast30time": mostlast30time,
		"leastlast30time": leastlast30time,
	}

	return render(request, "Post\\postViewMobile.html", context)

def getData(request):
	notEmp = ['Test_employee', 'Ajinkya Mukhedkar', 'ajiNkya', 'AjinkyaM', 'Arnab', 'Arnab', 'Ajinkya Mukhedkar']
	siteid = request.GET.get('sitecode')
	

	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	# print("----------------ActiveQR Code-------->>>", len(activeQR)) #activeQR,
	place_list = [i.name for i in activeQR]
	# print("----------Place Name-->>>", place_list)
	uuidList = []
	for place in place_list:	
		qrList = uuidTbl.objects.get(placeid_id=place)
		# print("qrList-->>", qrList.uuid)
		uuidList.append(qrList.uuid)
	monitor = []
	lastVisit = []
	lastVisitShort = []
	hrsLastVisit = []
	endHrs = datetime.now().strftime('%d %m %y  %H:%M %p')
	endDHrs = datetime.strptime(endHrs, '%d %m %y  %H:%M %p')

	startHrs = (endDHrs - timedelta(hours=24)).strftime('%d %m %y  %H:%M %p')
	startDHrs = datetime.strptime(startHrs, '%d %m %y  %H:%M %p')
	for qrName in uuidList:
		start = (date.today() - timedelta(days=1)).strftime('%Y %m %d %H:%M:%S')
		startD = datetime.strptime(start, '%Y %m %d %H:%M:%S')
		end = datetime.now().strftime('%Y %m %d %H:%M:%S')
		endD = datetime.strptime(end, '%Y %m %d %H:%M:%S')
		# print(f"endD--->>>{endD}---startD--->>>{startD}")
		entry = pathTbl.objects.filter(time__range = (startD, endD), uuid=qrName)
		# print(qrName, "Monitored----------------->>>", len(entry))
		monitor.append(len(entry))
		visitTime = []
		for vis in entry:
			visitTime.append(vis.time)
		if not len(entry):
			entry = pathTbl.objects.filter(time__lte = endD, uuid=qrName)
			for vis in entry:
				visitTime.append(vis.time)
		if visitTime:
			lastVisit.append(visitTime[-1].strftime("%A %d-%b-%Y %H:%M:%S"))
			# lastVisitShort.append(visitTime[-1])
			lastVisitShort.append(visitTime[-1].strftime("%a %d %b %y %I:%M %p"))
		else:
			lastVisit.append('0')
		
		visitTime.clear()
	most = max(monitor)
	least = min(monitor)
	# print("-----------------------------------------------------------------")
	# print(f"Most--->>>{most}----Least--->>>{least}")
	# print("UUID List Name--->>>", uuidList)
	# print("Scanned Counts--->>>", monitor)
	mostScanned = uuidTbl.objects.get(uuid=uuidList[monitor.index(most)]).placeid
	mostScannedUuid = uuidTbl.objects.get(uuid=uuidList[monitor.index(most)]).uuid

	leastScanned = uuidTbl.objects.get(uuid=uuidList[monitor.index(least)]).placeid
	leastScannedUuid = uuidTbl.objects.get(uuid=uuidList[monitor.index(least)]).uuid
	# print(f"Most Scanned--->>>{mostScanned}----Least Scanned--->>>{leastScanned}")

	time_list = [(8, 9), (10, 11), (12, 13), (14, 15), (16, 17), (18, 19), (20, 21), (22, 23), (0, 1), (2, 3), (4, 5), (6, 7)]
	# splitTimeMonitorList = []
	count = 0
	# for i in time_list:
		# print("Time split is--->>>", i)
		# splitTimeStart = datetime(2022, 7, 26, i[0], 00, 00)
		# splitTimeEnd = datetime(2022, 7, 26, i[1], 59, 59)
		# print("Time split is--->>>", i, "splitTimeStart--->>>", splitTimeStart, "splitTimeEnd--->>>", splitTimeEnd)
		# postData = pathTbl.objects.filter(time__gte = splitTimeStart, time__lte = splitTimeEnd)
		# splitTimeMonitorList.append(len(postData))
		# print(f"--------------Start--->>>{splitTimeStart}--------------End--->>>{splitTimeEnd}----------")
		# print("Split Time Data--->>>", postData, "Len--->>>", len(postData))
		# count+=len(postData)

	# 30 days data
	# print("------------------Date--------------------")
	filterFrom = request.GET.get('date_inp').split('-')
	filterTo = request.GET.get('date_inp1').split('-')
	name = request.GET.get('name')
	postName = request.GET.get('postcode')
	print("filterFrom------->>>>", filterFrom, "filterTo--->>>", filterTo)
	print("Sitecode is--->>>", siteid)
	print("EmpName is--->>>", name)

	start_dt = date(int(filterFrom[0]), int(filterFrom[1]), int(filterFrom[2]))
	end_dt = date(int(filterTo[0]), int(filterTo[1]), int(filterTo[2]))
	last30 = start_dt.strftime("%Y %m %d") + " to " + end_dt.strftime("%Y %m %d")
	oneMonthList = []
	date_list = []

	# print("------Start Dt--->>>", start_dt)
	# print("------End Dt--->>>", end_dt)
	for dt in daterange(start_dt, end_dt):
		date_list.append(dt.strftime("%Y %m %d"))

	# print(date_list)
	for i in time_list:
		date_data = []
		for dt in date_list:
			# print("--------Dt---------", dt)
			dateof = dt.split()
			print("--------------int(dateof[0]), int(dateof[1]), int(dateof[2])----------")
			print(int(dateof[0]), int(dateof[1]), int(dateof[2]))
			splitTimeStartDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[0], 00, 00)
			splitTimeEndDay = datetime(int(dateof[0]), int(dateof[1]), int(dateof[2]), i[1], 59, 59)
			if name:
				empid = empTbl.objects.get(name=name)
				postData = pathTbl.objects.filter(time__gte = splitTimeStartDay, time__lte = splitTimeEndDay, empid_id=empid)
				date_data.append(len(postData))
		
			elif postName:
				postid = uuidTbl.objects.get(placeid_id=postName)
				postData = pathTbl.objects.filter(time__gte = splitTimeStartDay, time__lte = splitTimeEndDay, uuid_id=postid)
				date_data.append(len(postData))
		
			else:
				postData = pathTbl.objects.filter(time__gte = splitTimeStartDay, time__lte = splitTimeEndDay)
				date_data.append(len(postData))
		# print("date_data-->>", date_data)
		oneMonthList.append(sum(date_data))
			# print(f"--------------Start--->>>{splitTimeStartDay}--------------End--->>>{splitTimeEndDay}----------")
			# print("Split Time Data--->>>", postData, "Len--->>>", len(postData))
	# print("-------------One Month List--------------")
	# print(oneMonthList)
	# print("-----------------------------------------")
	# print(time_list)
	# print(splitTimeMonitorList)
	# print(count)
	# siteList = list(placeTbl.objects.filter(siteid_id=siteid))
	uuid_name_list = list(zip(uuidList, place_list))
	# print("Mapped--->>>", uuid_name_list)
	empIdList = []
	empNameList = []
	empOneDayId = []
	empOneDayName = []
	uuidOneDay = []

	for uid in uuidList:
		dateNow = date.today()
		dateFrom = dateNow - timedelta(days=29)
		sD = date.today() - timedelta(days=1)
		eD = date.today()
		sdStrt = sD.strftime("%Y %m %d").split()
		edStrt = eD.strftime("%Y %m %d").split()
		from_D = datetime(int(sdStrt[0]), int(sdStrt[1]), int(sdStrt[2]), 20, 00, 00)
		to_D = datetime(int(edStrt[0]), int(edStrt[1]), int(edStrt[2]), 10, 00, 00)
		empId = pathTbl.objects.filter(uuid=uid, time__range = (dateFrom, dateNow))
		empIdOneDay = pathTbl.objects.filter(uuid=uid, time__range = (from_D, to_D))
		# print("empIdOneDay--------------", empIdOneDay)
		for emp in empId:
			if emp.empid_id not in empIdList:
				empIdList.append(emp.empid_id)
		if empIdOneDay:
			uuidOneDay.append(uid)
			for emp in empIdOneDay:
				if emp.empid_id not in empOneDayId:
					empOneDayId.append(emp.empid_id)

	for emp in empIdList:
		empName = empTbl.objects.get(empid=emp).name
		empNameList.append(empName)

	for emp in empOneDayId:
		empName = empTbl.objects.get(empid=emp).name
		empOneDayName.append(empName)
	# print("Employee List--->>>", empIdList)
	# print("Emp Name List--->>>", empNameList)
	empid_name_list = list(zip(empIdList, empNameList))

	print("empid_name_list before----->>>", empid_name_list)
	for i in empid_name_list:
		if i[1] in notEmp:
			empid_name_list.remove(i)
	print("empid_name_list After----->>>", empid_name_list)

	total = sum(monitor)
	# print("Total---",total)
	
	visitPer = [round((int(i)/int(total))*100, 2) for i in monitor]
	
	monthDict = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07',
	 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
	for i in lastVisit:
		today = datetime.now().strftime('%Y %d %m %H:%M:%S')
		todayDT = datetime.strptime(today, '%Y %d %m %H:%M:%S')
		print(today)
		if i != '0':
			val = i.split()
			dateVal = val[1].split('-')
			timeVal = val[2].split(':')
			final = dateVal + timeVal
			visitVal = f"{final[2]} {final[0]} {monthDict.get(final[1])} {final[3]}:{final[4]}:{final[5]}"
			print(final)
			visitValDT = datetime.strptime(visitVal, '%Y %d %m %H:%M:%S')
			print(type(visitValDT), visitValDT)
			diff = todayDT-visitValDT
			diffSec = diff.total_seconds()
			difHours = round(diffSec)/3600
			hrsLastVisit.append(math.ceil(difHours))
		else:
			hrsLastVisit.append('0')

	# print("Site List-->>",uuidList)
	# print("Site Name List-->>",place_list)
	# print("Monitor--->>>", monitor)
	# print("Visit Percentage--->>>", visitPer)
	# print("Last Visit--->>>", lastVisit)
	# print("Since--->>>", hrsLastVisit)
	OneDayData = list(zip(empNameList, place_list, monitor, visitPer, lastVisitShort, hrsLastVisit))
	sD = date.today() - timedelta(days=1)
	eD = date.today()
	sdStrt = sD.strftime("%Y %m %d").split()
	edStrt = eD.strftime("%Y %m %d").split()
	from_D = datetime(int(sdStrt[0]), int(sdStrt[1]), int(sdStrt[2]), 20, 00, 00)
	to_D = datetime(int(edStrt[0]), int(edStrt[1]), int(edStrt[2]), 10, 00, 00)
	scanned = []
	# print("empIdList", empIdList)
	# print("empNameList", empNameList)
	for i in empOneDayId:
		empWiseData = pathTbl.objects.filter(time__gte = from_D, time__lte = to_D, empid_id=i)
		# print("empWiseData--------------------------", empWiseData)
		if len(empWiseData):
			scanned.append(len(empWiseData))
		else:
			scanned.append(0)
		# print("Scanning data----->>>>", i, len(empWiseData))
	# print("Scanned List-->>", scanned)
	most_scanned = max(scanned)
	least_scanned = min(scanned)
	most_scanned_by = empOneDayName[scanned.index(most_scanned)]
	most_scanned_by_empId = empOneDayId[scanned.index(most_scanned)]
	
	least_scanned_by = empOneDayName[scanned.index(least_scanned)]
	least_scanned_by_empId = empOneDayId[scanned.index(least_scanned)]
	totalPostMonitored = sum(oneMonthList)
	# print("-------------------------------OneDayData---------------------------")
	# print(OneDayData)
	# print("----------------------------------END--------------------------------")

	mostof30 = max(oneMonthList)
	leastof30 = min(oneMonthList)
	time_val_list = ['8AM to 10AM','10AM to 12NOON','12NOON to 2PM','2PM to 4PM','4PM to 6PM','6PM to 8PM','8PM to 10PM','10PM to 00Hour','00HOUR to 2AM','2AM to 4AM','4AM to 6AM','6AM to 8AM']
	mostlast30time = time_val_list[oneMonthList.index(mostof30)]
	leastlast30time = time_val_list[oneMonthList.index(leastof30)]

	placeNameList = []
	for i in place_list:
		ind = i.index(',')
		placeNameList.append(i[:ind])
	print(placeNameList, monitor)

	lastVisitQR = []
	hrsLastVisitQR = []
	for qrName in uuidList:
		upto = datetime.now()
		entry = pathTbl.objects.filter(time__lte = upto, uuid=qrName)
		visitTime = []
		for vis in entry:
			visitTime.append(vis.time)
		if visitTime:
			lastVisitQR.append(visitTime[-1].strftime("%Y %m %d %H:%M:%S"))
	print("lastVisitQR", lastVisitQR)
	calcDiff = list(zip(lastVisitQR,place_list))
	differenceHours = []

	for lastVis in calcDiff:
		today = datetime.now().strftime('%Y %m %d %H:%M:%S')
		todayDT = datetime.strptime(today, '%Y %m %d %H:%M:%S')
		lastvis = datetime.strptime(lastVis[0], '%Y %m %d %H:%M:%S')
		diff = todayDT - lastvis
		diffSec = diff.total_seconds()
		difHours = round(diffSec)/3600
		# print("diffSec", diffSec)
		# print("difHours", difHours)
		if difHours > 2:
			differenceHours.append([lastVis[1], diff, math.ceil(difHours)])
	finaldifferenceHours = sorted(differenceHours, key=lambda x: x[-1], reverse=True)
	# print("finaldifferenceHours---", finaldifferenceHours)
	context = {
		# 'data' : splitTimeMonitorList,
		'data' : oneMonthList,
		'start':start_dt,
		'end': end_dt,
		'totalDays': len(date_list),
		'total': totalPostMonitored,
		'last30': last30,
		"mostlast30time": mostlast30time,
		"leastlast30time": leastlast30time,
		'sitecode' : siteid,
		'activePost' : len(activeQR),
		'mostScanned' : mostScanned,
		'leastScanned' : leastScanned,
		"postid": postName,
		"name": name,
		'uuidList' : uuid_name_list,
		'empdata' : empid_name_list,
		'oneDayData' : OneDayData,
		'fromD' : from_D,
		"toD" : to_D,
		"dateFrom" : startDHrs,
		"dateTo" : endDHrs,
		"most_scanned_by" : most_scanned_by,
		'most_scanned_by_empId': most_scanned_by_empId,
		"least_scanned_by" : least_scanned_by,
		"least_scanned_by_empId" : least_scanned_by_empId,
		"buildingData": monitor,
		"buildingName": placeNameList,
		'most_scanned': most_scanned,
		'least_scanned': least_scanned,
		'mostScannedUuid': mostScannedUuid,
		'leastScannedUuid': leastScannedUuid,
	}

	if differenceHours:
		context['differenceHours'] = finaldifferenceHours

	if name:
		return render(request, "Post\\empViewMobile.html", context)
	elif postName:
		return render(request, "Post\\postViewMobile.html", context)
	else:
		return render(request, "Post\\qrDashboardMobile.html", context)

def daterange(date1, date2):
	print("called-------------")
	for n in range(int ((date2 - date1).days)+1):
		yield date1 + timedelta(n)

def getApiData(request):
	url = "https://e-attendance.starspayroll.in/api/external/site-duty-roaster?site_code=FCM/KOL/INV/SRWD/01&roaster_date=2022-08-18"
	api_data = requests.get(url, params=request.GET).text
	info = json.loads(api_data)

	site_code_list = []
	site_name_list = []
	category_code_list = []
	category_name_list = []
	shift_code_list = []
	shift_hours_list = []
	employee_code_list = []
	employee_name_list = []
	from_date_list = []
	withdrawn_date_list = []
	allocation_type_list = []
	roaster_type_list = []
	off_day_list = []
	roaming_employee_list = []

	for val in info.get('data'):
		site_code_list.append(val.get('site_code'))
		site_name_list.append(val.get('site_name'))
		category_code_list.append(val.get('category_code'))
		category_name_list.append(val.get('category_name'))
		shift_code_list.append(val.get('shift_code'))
		shift_hours_list.append(val.get('shift_hours'))
		employee_code_list.append(val.get('employee_code'))
		employee_name_list.append(val.get('employee_name'))
		from_date_list.append(val.get('from_date'))
		withdrawn_date_list.append(val.get('withdrawn_date'))
		allocation_type_list.append(val.get('allocation_type'))
		roaster_type_list.append(val.get('roaster_type'))
		off_day_list.append(val.get('off_day'))
		roaming_employee_list.append(val.get('roaming_employee'))
	res = {
		'site_code': site_code_list,
		'site_name': site_name_list,
		'category_code': category_code_list,
		'category_name': category_name_list,
		'shift_code': shift_code_list,
		'shift_hours': shift_hours_list,
		'employee_code': employee_code_list,
		'employee_name': employee_name_list,
		'from_date': from_date_list,
		'withdrawn_date': withdrawn_date_list,
		'allocation_type': allocation_type_list,
		'roaster_type': roaster_type_list,
		'off_day': off_day_list,
		'roaming_employee': roaming_employee_list
	}

	final = pd.DataFrame.from_dict(res)
	fileNameApiData = 'Api Data by Shift.xlsx'
	final.to_excel(fileNameApiData)
	return HttpResponse('Done..........')

def qrReport(request):
	siteid = request.GET.get('sitecode')
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	# print("Place List--->>>", place_list)
	# place_list_siteid = [i.name+' ('+siteid+')' for i in activeQR]
	
	
	allPlace = []
	uuidList = []
	place_list_siteid = []
	lenSite = len(siteid)
	place_list_siteid_1 = [i.replace(f'({siteid})', '') for i in place_list]
	place_list_siteid_short = []
	for place in place_list:	
		qrList = uuidTbl.objects.filter(placeid_id=place)
		for placeuid in qrList:
			place_list_siteid.append("( "+placeuid.name+" ) "+place)
			allPlace.append(placeuid.name)
			uuidList.append(placeuid.uuid)
	# print("allPlace--->>>", allPlace , len(allPlace))
	# postNo = []
	# for i in range(1, len(place_list_siteid)+1):
	# 	postNo.append('Post No.'+str(i))
	lenList = [len(i) for i in place_list_siteid]
	# One Day Post Monitoring Report 10 to 10
	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()

	dt_1=yest.strftime("%d-%b-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	shiftA = [(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00), (9,00,9,30), (9,30,10,00),
			(10,00,10,30), (10,30,11,00), (11,00,11,30), (11,30,12,00), (12,00,12,30),
			(12,30,13,00), (13,00,13,30), (13,30,14,00)]

	shiftB = [(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
			(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00),
			(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00)]

	shiftC = [(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
			(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), (2,00,2,30),
			(2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30),
			(5,30,6,00), (6,00,6,30), (6,30,7,00)]

	# shiftA = [(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00), (9,00,9,30), (9,30,10,00),
	# 		(10,00,10,30), (10,30,11,00), (11,00,11,30), (11,30,12,00), (12,00,12,30),
	# 		(12,30,13,00), (13,00,13,30), (13,30,14,00)]

	# shiftB = [(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
	# 		(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00),
	# 		(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00)]

	# shiftC = [(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
	# 		(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), (2,00,2,30),
	# 		(2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30),
	# 		(5,30,6,00), (6,00,6,30), (6,30,7,00)]

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []
	tz=timezone.get_current_timezone()

	for qrid in uuidList:
		isQrData = []
		for t in shiftA:
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)

			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

	isDataShiftB = []
	AttendedShiftB = []
	unAttendedShiftB = []
	for qrid in uuidList:
		isQrData = []
		for t in shiftB:
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftB.append(isQrData.count(0))
		AttendedShiftB.append(isQrData.count(1))
		isDataShiftB.append(isQrData)

	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftB----->>>>>", isDataShiftB, "\n")

	isDataShiftC = []
	AttendedShiftC = []
	unAttendedShiftC = []
	for qrid in uuidList:
		isQrData = []
		for t in shiftC:
			if t == (23,30,00,00):
				# print("Inside If")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			elif shiftC.index(t) > shiftC.index((23,30,00,00)):
				# print("Inside Elif")
				# print(t)
				start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			else:
				# print("Else Block")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
		unAttendedShiftC.append(isQrData.count(0))
		AttendedShiftC.append(isQrData.count(1))
		isDataShiftC.append(isQrData)

	nightShift = [(21,00,23,00), (23,00,1,00), (1,00,3,00), (3,00,5,00), (5,00,7,00)]
	nShift = [('21:00','23:00'), ('23:00','1:00'), ('1:00','3:00'), ('3:00','5:00'), ('5:00','7:00')]
	isDataNightShift = []
	# AttendedNightShift = []
	# unAttendedNightShift = []

	for qrid in uuidList:
		isQrData = []
		for t in nightShift:
			if t == (23,00,1,00):
				# print("Inside If")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			elif nightShift.index(t) > nightShift.index((23,00,1,00)):
				# print("Inside Elif")
				# print(t)
				start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			else:
				# print("Else Block")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
		# unAttendedNightShift.append(isQrData.count(0))
		# AttendedNightShift.append(isQrData.count(1))
		isDataNightShift.append(isQrData)

	print("----------isDataNightShift--------")
	print(isDataNightShift)

	# print("AttendedShiftA", AttendedShiftA)
	# print("unAttendedShiftA", unAttendedShiftA)
	# print("AttendedShiftB", AttendedShiftB)
	# print("unAttendedShiftB", unAttendedShiftB)
	# print("AttendedShiftC", AttendedShiftC)
	# print("unAttendedShiftC", unAttendedShiftC)
	# print(place_list_siteid)

	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	maxAShiftC = max(AttendedShiftC)
	maxUShiftC = max(unAttendedShiftC)
	maxLen = int(max(lenList))

	module_dir = os.path.dirname(__file__)

	if siteid=="AMRI-Mukundpur":
		filename = os.path.join(module_dir,"Post Monitoring Report Format_time.xlsx")
	else:
		filename = os.path.join(module_dir,"Post Monitoring Report Format.xlsx")
	# filename ="Post Monitoring Report Format.xlsx"
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Shift A"]
	ws2 = wb2["Shift B"]
	ws3 = wb2["Shift C"]

	ws1['A4'] = siteid
	ws2['A4'] = siteid
	ws3['A4'] = siteid

	ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
	ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
	ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
	thin = Side (border_style="thin", color="000000")


	shiftAcol = 3
	shiftAi = 0
	for post in place_list_siteid:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1

	shiftBcol = 3
	shiftBi = 0
	for post in place_list_siteid:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	shiftCcol = 3
	shiftCi = 0
	for post in place_list_siteid:
		shiftCrow = 19
		for data in isDataShiftC[shiftCi]:
			# ws3.cell(row = shiftCrow, column = shiftCcol, value=post)
			# ws3.cell(row = shiftCrow, column = shiftCcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftCrow += 1
		shiftCi += 1
		shiftCcol += 1

	col = 3
	for post in range(0, len(place_list_siteid)):
		ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteid[post])
		
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")


		ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1['C17'] = 'Post Name'
	ws2['C17'] = 'Post Name'
	ws3['C17'] = 'Post Name'
	ws1['C17'].font = Font(bold=True)
	ws2['C17'].font = Font(bold=True)
	ws3['C17'].font = Font(bold=True)

	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteid)):
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	# Night shift data is commented as of now...
	# nshiftcol = 2
	# nshifti = 0
	# nshiftrow = 46
	# srNo = 1
	# for post in place_list_siteid:
	# 	tFrom = 0
	# 	tTo = 1
	# 	iTime = 0
	# 	for data in isDataNightShift[nshifti]:
	# 		if not data:

	# 			fCell = 'B'+str(nshiftrow)
	# 			tCell = 'I'+str(nshiftrow)
	# 			ws3.merge_cells(f'{fCell}:{tCell}')
	# 			ws3.cell(row = nshiftrow, column = 1, value=srNo)
	# 			ws3.cell(row = nshiftrow, column = 1).alignment = Alignment(horizontal='center', vertical='center')
	# 			ws3.cell(row = nshiftrow, column = nshiftcol, value=post)
	# 			ws3.cell(row = nshiftrow, column = 10, value=nShift[iTime][tFrom])
	# 			ws3.cell(row = nshiftrow, column = 11, value=nShift[iTime][tTo])
	# 			ws3.cell(row = nshiftrow, column = 10).alignment = Alignment(horizontal='center', vertical='center')
	# 			ws3.cell(row = nshiftrow, column = 11).alignment = Alignment(horizontal='center', vertical='center')
	# 			ws3.cell(row = nshiftrow, column = nshiftcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 			# ws1.column_dimensions[].auto_size = int(len(place_list_siteid[unAttendedShiftA.index(maxUShiftA)]))
				
	# 			nshiftrow += 1
	# 			srNo += 1
	# 		iTime += 1
	# 	nshifti += 1
	# Border for night shift table
	# thin = Side (border_style="thin", color="000000")
	# aCell = 'A'+str(46)
	# kCell = 'K'+str(nshiftrow-1)
	# for row in ws3[f'{aCell}:{kCell}']:
	# 	for cell in row:
	# 		cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# End of night shift

	ws1['E8'] = place_list_siteid[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteid[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteid[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteid[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteid[unAttendedShiftC.index(maxUShiftC)]
	ws3['E9'] = place_list_siteid[AttendedShiftC.index(maxAShiftC)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')

	# saving the destination excel file
	# wb2.save(str(filename1))

	if siteid=="AMRI-Mukundpur":

		ws4 = wb2["Timestamp"]

		time1=datetime.now()
		time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		time2_1=time2_1+timedelta(hours=7)
		time2_11=time2_1-timedelta(days=1)

		start_time1=tz.localize(time2_11)
		end_time1=tz.localize(time2_1)

		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=siteid,time__range=(start_time1,end_time1)).values()))

		if len(all_path)!=0:

			all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
			all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

			df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')


			all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
			all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

			df_all=pd.merge(df_all1,all_emp,on='empid_id')

			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			# df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
			df_all=df_all.sort_values("time",ascending=True)
			df_all=df_all.reset_index()

			# print(df_all.columns)

			rownum=2

			for idx,rows in df_all.iterrows():
				ws4.cell(row = rownum, column = 1, value=rows['name'])
				ws4.cell(row = rownum, column = 2, value=rows['placeid_id'])
				ws4.cell(row = rownum, column = 3, value=rows['datetime1'])
				ws4.cell(row = rownum, column = 4, value=rows['empname'])

				ws4.cell(row = rownum, column = 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws4.cell(row = rownum, column = 2).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws4.cell(row = rownum, column = 3).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws4.cell(row = rownum, column = 4).border = Border(top=thin, left=thin, right=thin, bottom=thin)

				rownum=rownum+1



			wb2.save(str(filename1))
		else:
			wb2.save(str(filename1))

	else:
		wb2.save(str(filename1))


	try:
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {siteid} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()




	return HttpResponse("Success")

def qrReportevent(request):
	siteid = request.GET.get('sitecode',"")
	atime=request.GET.get('atime',"")
	btime=request.GET.get('btime',"")
	ctime=request.GET.get('ctime',"")
	freq=request.GET.get('freq',"")
	onlynight=request.GET.get('night',"")

	removea=request.GET.get("removea","")
	removeb=request.GET.get("removeb","")

	if float(onlynight)!=1:
		if float(ctime)==0:
			atime=int(float(atime))
			btime=int(float(btime))

			shiftA=[]
			shiftB=[]

			if float(freq)==30:
				tot_slotA=(btime-atime)
				tot_slotB=((24-btime)+atime)
				atime1=atime
				btime1=btime

				i=0
				while i<tot_slotA:
					ele1=(atime1,00,atime1,30)
					ele2=(atime1,30,atime1+1,00)
					shiftA.append(ele1)
					shiftA.append(ele2)
					i=i+1
					atime1=atime1+1

				j=0

				while j<tot_slotB:
					if btime1!=23:
						ele1=(btime1,00,btime1,30)
						ele2=(btime1,30,btime1+1,00)
						btime1=btime1+1
					else:
						ele1=(btime1,00,btime1,30)
						ele2=(btime1,30,00,00)
						btime1=0

					shiftB.append(ele1)
					shiftB.append(ele2)
					j=j+1

			elif float(freq)==60:
				tot_slotA=(btime-atime)
				tot_slotB=((24-btime)+atime)
				atime1=atime
				btime1=btime

				i=0
				while i<tot_slotA:
					ele1=(atime1,00,atime1+1,00)
					shiftA.append(ele1)
					i=i+1
					atime1=atime1+1

				j=0

				while j<tot_slotB:
					if btime1!=23:
						ele1=(btime1,00,btime1+1,00)
						btime1=btime1+1
					else:
						ele1=(btime1,00,00,00)

						btime1=0

					shiftB.append(ele1)
					j=j+1

			print(shiftA)
			print(shiftB)

			timeshiftA=[]
			timeshiftB=[]

			for ele_a in shiftA:
				if ele_a[1]<10:
					b1="0"+str(ele_a[1])
				else:
					b1=str(ele_a[1])

				if ele_a[3]<10:
					d1="0"+str(ele_a[3])
				else:
					d1=str(ele_a[3])

				timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

			for ele_a in shiftB:
				if ele_a[1]<10:
					b1="0"+str(ele_a[1])
				else:
					b1=str(ele_a[1])

				if ele_a[3]<10:
					d1="0"+str(ele_a[3])
				else:
					d1=str(ele_a[3])
				timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)


			if removea=="1":
				df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))
				df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))
				# print("program is here")
			elif removeb=="1":

				df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
				df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
			else:
				df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A")|Q(other="AB"),placeid__siteid__name=siteid).values()))
				df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B")|Q(other="AB"),placeid__siteid__name=siteid).values()))


			uuidListA=df_A["uuid"].tolist()
			df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
			place_list_siteidA=df_A["fullname"].tolist()

			uuidListB=df_B["uuid"].tolist()
			df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
			place_list_siteidB=df_B["fullname"].tolist()

			todayStr = date.today().strftime('%Y %d %m').split()
			today = date.today()
			yest = today - timedelta(days = 1)
			yestStr = yest.strftime('%Y %d %m').split()
			dt_1=yest.strftime("%d-%B-%Y")

			isDataShiftA = []
			AttendedShiftA = []
			unAttendedShiftA = []

			tz=timezone.get_current_timezone()

			for qrid in uuidListA:
				isQrData = []
				for t in shiftA:
					# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				unAttendedShiftA.append(isQrData.count(0))
				AttendedShiftA.append(isQrData.count(1))
				isDataShiftA.append(isQrData)


			if float(freq)==30:
				isDataShiftB = []
				AttendedShiftB = []
				unAttendedShiftB = []

				for qrid in uuidListB:

					isQrData = []
					for t in shiftB:
						if t == (23,30,00,00):
							# print("Inside If")
							# print(t)
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)
						elif shiftB.index(t) > shiftB.index((23,30,00,00)):
							# print("Inside Elif")
							# print(t)
							start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)
						else:
							# print("Else Block")
							# print(t)
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
							# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)

					unAttendedShiftB.append(isQrData.count(0))
					AttendedShiftB.append(isQrData.count(1))
					isDataShiftB.append(isQrData)

			elif float(freq)==60:

				isDataShiftB = []
				AttendedShiftB = []
				unAttendedShiftB = []

				for qrid in uuidListB:

					isQrData = []
					for t in shiftB:
						if t == (23,00,00,00):
							# print("Inside If")
							# print(t)
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)
						elif shiftB.index(t) > shiftB.index((23,00,00,00)):
							# print("Inside Elif")
							# print(t)
							start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)
						else:
							# print("Else Block")
							# print(t)
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
							# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isQrData.append(1)
							else:
								isQrData.append(0)

					unAttendedShiftB.append(isQrData.count(0))
					AttendedShiftB.append(isQrData.count(1))
					isDataShiftB.append(isQrData)

			else:
				pass


			maxAShiftA = max(AttendedShiftA)
			maxUShiftA = max(unAttendedShiftA)
			maxAShiftB = max(AttendedShiftB)
			maxUShiftB = max(unAttendedShiftB)

			module_dir = os.path.dirname(__file__)


			if siteid!="Himadri Chemical Ltd.":
				filename=os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime1.xlsx")
			else:
				filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime1.xlsx")


			wb1 = xl.load_workbook(filename)
			  
			# opening the destination excel file 
			if siteid!="Apollo Hospitals":
				filename1 =siteid+" Post Monitoring Report.xlsx"
			else:
				filename1 =siteid+" Patrol Monitoring Report.xlsx"
			wb2 = xl.load_workbook(filename)
			# ws2 = wb2.active
			ws1 = wb2["Shift A"]
			ws2 = wb2["Shift B"]
			# ws3 = wb2["Shift C"]
			# ws4 = wb2["Shift G"]
			ws1['A4'] = siteid
			ws2['A4'] = siteid
			# ws3['A4'] = siteid
			# ws4['A4'] = siteid
			ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" "+str(int(float(atime)))+":00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" "+str(int(float(atime)))+":00 Hours"
			ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" "+str(int(float(btime)))+":00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" "+str(int(float(atime)))+":00 Hours"


			rownum1=19
			thin = Side (border_style="thin", color="000000")

			for ele_ta in timeshiftA:
				ele_ta1=ele_ta.split("-")
				ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
				ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
				ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
				ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
				rownum1=rownum1+1



			shiftAcol = 3
			shiftAi = 0
			for postA in place_list_siteidA:
				shiftArow = 19
				for data in isDataShiftA[shiftAi]:
					# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
					# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
					if data:
						ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
						ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					else:
						ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
						ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
						
					shiftArow += 1
				shiftAi += 1
				shiftAcol += 1


			rownum2=19
			for ele_tb in timeshiftB:
				ele_tb1=ele_tb.split("-")
				ws2.cell(row=rownum2,column=1).value=ele_tb1[0]
				ws2.cell(row=rownum2,column=2).value=ele_tb1[1]
				ws2.cell(row=rownum2,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
				ws2.cell(row=rownum2,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
				rownum2=rownum2+1

			shiftBcol = 3
			shiftBi = 0
			for postB in place_list_siteidB:
				shiftBrow = 19
				for data in isDataShiftB[shiftBi]:
					# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
					# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
					if data:
						ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
						ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					else:
						ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
						ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					shiftBrow += 1
				shiftBi += 1
				shiftBcol += 1


			

			col = 3
			for post in range(0, len(place_list_siteidA)):
				# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
				ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
				
				# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				
				# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

				# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
				ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

				# ws1.column_dimensions[get_column_letter(col)].width = 25
				ws1.column_dimensions[get_column_letter(col)].width = 25
				# ws1.row_dimensions[18].height = 70
				ws1.row_dimensions[18].height = 70
				col += 1

			# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
			ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

			ws1['C17'] = 'Post Name'

			ws1['C17'].font = Font(bold=True)

			# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
			ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
			# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

			bordercol = 3
			for post in range(0, len(place_list_siteidA)):
				# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				bordercol += 1


			col = 3
			for post in range(0, len(place_list_siteidB)):
				# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
				ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
				
				# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				
				# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

				# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
				ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

				# ws1.column_dimensions[get_column_letter(col)].width = 25
				ws2.column_dimensions[get_column_letter(col)].width = 25
				# ws1.row_dimensions[18].height = 70
				ws2.row_dimensions[18].height = 70
				col += 1

			# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
			ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

			ws2['C17'] = 'Post Name'

			ws2['C17'].font = Font(bold=True)

			# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
			ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
			# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

			bordercol = 3
			for post in range(0, len(place_list_siteidB)):
				# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				bordercol += 1



			ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
			ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
			ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
			ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
			# ws1.column_dimensions['F'].auto_size = True

			ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
			ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
			ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
			ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

			# ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
			# ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
			# ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
			# ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

			# ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
			# ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
			# ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
			# ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

			ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
			ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
			ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
			ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')

			ws3 = wb2["Timestamp"]

			time1=datetime.now()
			time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			time2_1=time2_1+timedelta(hours=int(float(atime)))
			time2_11=time2_1-timedelta(days=1)

			start_time1=tz.localize(time2_11)
			end_time1=tz.localize(time2_1)

			all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=siteid,time__range=(start_time1,end_time1)).values()))

			if len(all_path)!=0:

				all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
				all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

				df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')


				all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
				all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

				df_all=pd.merge(df_all1,all_emp,on='empid_id')

				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

				# df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
				df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
				df_all=df_all.sort_values("time",ascending=True)
				df_all=df_all.reset_index()

				# print(df_all.columns)

				rownum=2

				for idx,rows in df_all.iterrows():
					ws3.cell(row = rownum, column = 1, value=rows['name'])
					ws3.cell(row = rownum, column = 2, value=rows['placeid_id'])
					ws3.cell(row = rownum, column = 3, value=rows['datetime1'])
					ws3.cell(row = rownum, column = 4, value=rows['empname'])

					ws3.cell(row = rownum, column = 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					ws3.cell(row = rownum, column = 2).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					ws3.cell(row = rownum, column = 3).border = Border(top=thin, left=thin, right=thin, bottom=thin)
					ws3.cell(row = rownum, column = 4).border = Border(top=thin, left=thin, right=thin, bottom=thin)

					rownum=rownum+1


			if removea=="1":
				wb2.remove(wb2['Shift A'])
			elif removeb=="1":
				wb2.remove(wb2['Shift B'])
			else:
				pass




			wb2.save(str(filename1))

	if siteid=="Cable TV ETMA":
		sitenamemail="CABLE TV EQUIPMENTS TRADERS & MANUFACTURERS ASSOCIATION"
	else:
		sitenamemail=siteid



	try:
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()









	return HttpResponse("Success")









def qrReportTwoShift(request):
	siteid = request.GET.get('sitecode')
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	# print("Place List--->>>", place_list)
	# place_list_siteid = [i.name+' ('+siteid+')' for i in activeQR]

	# notinlist=["G"]
	
	
	allPlace = []
	uuidList = []
	place_list_siteid = []
	lenSite = len(siteid)
	place_list_siteid_1 = [i.replace(f'({siteid})', '') for i in place_list]
	place_list_siteid_short = []

	if siteid=="Teloijan-Tea":
		for place in place_list:	
			qrList = uuidTbl.objects.filter(placeid_id=place)
			for placeuid in qrList:
				place_list_siteid.append("( "+placeuid.name+" ) "+place)
				allPlace.append(placeuid.name)
				uuidList.append(placeuid.uuid)
	elif siteid=="Vikram Solar":
		for place in place_list:	
			qrList = uuidTbl.objects.filter(placeid_id=place)
			for placeuid in qrList:
				if placeuid.name!="Gate-2":
					place_list_siteid.append("( "+placeuid.name+" ) "+place)
					allPlace.append(placeuid.name)
					uuidList.append(placeuid.uuid)
	else:
		for place in place_list:	
			qrList = uuidTbl.objects.filter(placeid_id=place)
			for placeuid in qrList:
				place_list_siteid.append("( "+placeuid.name+" ) "+place)
				allPlace.append(placeuid.name)
				uuidList.append(placeuid.uuid)



	# print("allPlace--->>>", allPlace , len(allPlace))
	# postNo = []
	# for i in range(1, len(place_list_siteid)+1):
	# 	postNo.append('Post No.'+str(i))
	lenList = [len(i) for i in place_list_siteid]
	# One Day Post Monitoring Report 10 to 10
	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	if siteid=="Teloijan-Tea" or siteid=="Tirumala Balaji Alloys" or siteid==" Ekkya House E006 11" or siteid=="Emami House E006 03" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="Astral Flat Owners ASAF02" or siteid=="SE Builders And Realtors UKAC02" or siteid=="SE Builders And Realtors UKAC01":

		shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	elif siteid=="Vikram Solar" or siteid=="Shree Mangal Trust" or siteid=="Merlin Projects Ltd" or siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16" or siteid=="South Point High School SPSC01":
		shiftB = [(22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]

	elif siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA" or siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="EMAMI REALTY LTD EIL201":
		shiftB = [(22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]

	elif siteid=="TRIPURA ENCLAVE TRPR01" or siteid=="ACC CEMENT ACCL01":
		shiftB = [(23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]
	elif siteid=="Pioneer Packaging PPCI02":
		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]


	elif siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="JJ Gold Refinery JGLD01" or siteid=="Sanjay Creation SCIN 01" or siteid=="Tata Asset Management TTAL01" or siteid=="Randstad India M012 01" or siteid=="BD Memorial School BDML05" or siteid=="Pulse Diagnostics PLSE03":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	elif siteid=="City Bazaar Metro CIBM01" or siteid=="BD Memorial School BDML04" or siteid=="Wow Momo Foods SIL MOMO01":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00)]


	elif siteid=="Cygnet Industries Ltd KSIL01":
		shiftA = [(10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00)]

	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06":
		shiftA=[(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00)]
	elif siteid=="Spencer Retail SPRL07":
		shiftA=[(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00)]
	elif siteid=="Spencer Retail SPRL67":
		shiftA=[(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00),(22,00,23,00)]




	elif siteid=="MUMUSO Retail MSLD 01" or siteid=="Fortune Marketing FMTG01":
		shiftA = [(10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00)]

	elif siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Chandrani Enterprises KOL CNDR02":
		shiftA = [(9,00,10,00),(10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00)]

	elif siteid=="Future Merchandising FVRL12" or siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Chandrani Enterprises KOL CNDR01":
		shiftA = [(9,00,10,00),(10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00)]
	elif siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00)]





	elif siteid=="YOKOGAWA INDIA LIMITED":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00)]
	elif siteid=="MK Enterprise":
		shiftA = [(10,00,10,30), (10,30,11,00),
			(11,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00),(14,00,14,30),(14,30,15,00),(15,00,15,30),(15,30,16,00),(16,00,16,30),(16,30,17,00),(17,00,17,30),
			(17,30,18,00),(18,00,18,30),(18,30,19,00),(19,00,19,30),(19,30,20,00),(20,00,20,30),(20,30,21,00),(21,00,21,30)]

	elif siteid=="Future Merchandising FVRL05" or siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01":
		shiftA = [(10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00),]


	elif siteid=="MK MKEE 03":
		shiftA = [(10,00,10,30), (10,30,11,00),
			(11,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00),(14,00,14,30),(14,30,15,00),(15,00,15,30),(15,30,16,00),(16,00,16,30),(16,30,17,00),(17,00,17,30),
			(17,30,18,00),(18,00,18,30),(18,30,19,00),(19,00,19,30),(19,30,20,00),(20,00,20,30),(20,30,21,00),(21,00,21,30)]



	elif siteid=="Berger Paints":
		shiftB=[(20,30,21,30), (21,30,22,30), (22,30,23,30), (23,30,00,30),
					(00,30,1,30), (1,30,2,30), (2,30,3,30), (3,30,4,30),
					(4,30,5,30), (5,30,6,30),]
	elif siteid=="Anjali Jewellers ANJL 03":
		shiftB=[(21,00,21,30), (21,30,22,00),(22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00),(8,00,8,30),(8,30,9,00)]
	elif siteid=="Anjali Jewellers ANJL 04":
		shiftB=[(20,00,20,30),(20,30,21,00),(21,00,21,30), (21,30,22,00),(22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00)]



	elif siteid=="Hiland Sapphire" or siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical":
		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00),(6,00,7,00)]

	elif siteid=="Rameswaram Mansion":
		shiftB = [(22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00)]
	elif siteid=="Narayan Hospital":
		shiftB = [(23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]

	elif siteid=="Calcutta Lions Netra Niketan CLSN00001":
		shiftB = [(19,00,20,00),(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00), (6,00,7,00)]

	elif siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00),(22,00,23,00)]





	elif siteid=="Veekay Industries":
		shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00)]

	elif siteid=="Chandrani Enterprises CNDR03" or siteid=="Emami Limited GUW EMMI06":
		shiftA = [(9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
			(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00),(20,00,21,00)]





	else:
		shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
					(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
					(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]


	timeshiftA=[]
	timeshiftB=[]

	if siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="YOKOGAWA INDIA LIMITED" or siteid=="MK Enterprise" or siteid=="MK MKEE 03" or siteid=="Cygnet Industries Ltd KSIL01" or siteid=="Future Merchandising FVRL12" or siteid=="Randstad India M012 01":

		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])


			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	elif siteid=="JJ Gold Refinery JGLD01" or siteid=="MUMUSO Retail MSLD 01" or siteid=="City Bazaar Metro CIBM01" or siteid=="Sanjay Creation SCIN 01" or siteid=="Future Merchandising FVRL05" or siteid=="Tata Asset Management TTAL01" or siteid=="Aeropict Logistics Solutions ARLG01":
		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])


			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	elif siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Fortune Marketing FMTG01" or siteid=="BD Memorial School BDML04" or siteid=="BD Memorial School BDML05" or siteid=="Chandrani Enterprises KOL CNDR01" or siteid=="Chandrani Enterprises KOL CNDR02" or siteid=="Pulse Diagnostics PLSE03" or siteid=="Wow Momo Foods SIL MOMO01":
		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])


			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06" or siteid=="Spencer Retail SPRL07" or siteid=="Spencer Retail SPRL67":
		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])


			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)


	elif siteid=="Chandrani Enterprises CNDR03" or siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04" or siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01" or siteid=="Emami Limited GUW EMMI06" or siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63":
		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])


			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)



	elif siteid=="South Point High School SPSC01" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA" or siteid=="TRIPURA ENCLAVE TRPR01" or siteid==siteid=="Anjali Jewellers ANJL 03":
		for ele_a in shiftB:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)



	elif siteid=="Berger Paints" or siteid=="Shree Mangal Trust" or siteid=="Merlin Projects Ltd" or siteid=="Hiland Sapphire" or siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical":
		for ele_a in shiftB:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	elif siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16" or siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="Astral Flat Owners ASAF02":
		for ele_a in shiftB:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)



	elif siteid=="Rameswaram Mansion" or siteid=="Narayan Hospital" or siteid=="Veekay Industries" or siteid=="Sundrex Oil Company Ltd SOCL01" or siteid=="Calcutta Lions Netra Niketan CLSN00001":
		for ele_a in shiftB:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	elif siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="ACC CEMENT ACCL01" or siteid=="Pioneer Packaging PPCI02" or siteid=="EMAMI REALTY LTD EIL201" or siteid=="SE Builders And Realtors UKAC01" or siteid=="SE Builders And Realtors UKAC02" or siteid=="Anjali Jewellers ANJL 04":
		for ele_a in shiftB:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftB.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)



	else:
		pass



	tz=timezone.get_current_timezone()






	if siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="YOKOGAWA INDIA LIMITED" or siteid=="MK Enterprise" or siteid=="MK MKEE 03" or siteid=="Cygnet Industries Ltd KSIL01" or siteid=="JJ Gold Refinery JGLD01" or siteid=="MUMUSO Retail MSLD 01":


		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		

		for qrid in uuidList:
			isQrData = []
			for t in shiftA:
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)

	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06" or siteid=="Spencer Retail SPRL07" or siteid=="Spencer Retail SPRL67":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		

		for qrid in uuidList:
			isQrData = []
			for t in shiftA:
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)


	elif siteid=="Sanjay Creation SCIN 01" or siteid=="City Bazaar Metro CIBM01" or siteid=="Future Merchandising FVRL05" or siteid=="Tata Asset Management TTAL01" or siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Future Merchandising FVRL12" or siteid=="Randstad India M012 01" or siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Fortune Marketing FMTG01" or siteid=="BD Memorial School BDML04":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		

		for qrid in uuidList:
			isQrData = []
			for t in shiftA:
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)

	elif siteid=="Chandrani Enterprises KOL CNDR01" or siteid=="BD Memorial School BDML05" or siteid=="Chandrani Enterprises KOL CNDR02" or siteid=="Pulse Diagnostics PLSE03" or siteid=="Chandrani Enterprises CNDR03" or siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		

		for qrid in uuidList:
			isQrData = []
			for t in shiftA:
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)

	elif siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01" or siteid=="Emami Limited GUW EMMI06" or siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63" or siteid=="Wow Momo Foods SIL MOMO01":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		

		for qrid in uuidList:
			isQrData = []
			for t in shiftA:
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)






	elif siteid=="Berger Paints":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		for qrid in uuidList:
			isQrData = []
			for t in shiftB:
				if t == (23,30,00,30):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,30,00,30)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)

	
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)

	elif siteid=="Anjali Jewellers ANJL 03" or siteid=="Anjali Jewellers ANJL 04":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		for qrid in uuidList:
			isQrData = []
			for t in shiftB:
				if t == (23,30,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,30,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)

	
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)



	else:

		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []
		for qrid in uuidList:
			isQrData = []
			for t in shiftB:
				if t == (23,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)

	
		maxAShiftB = max(AttendedShiftB)
		maxUShiftB = max(unAttendedShiftB)
	
	maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)
	if siteid=="Teloijan-Tea":
		filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift (1).xlsx")
	elif siteid=="Vikram Solar":
		filename = os.path.join(module_dir,"PMR2Sv.xlsx")
	elif siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="YOKOGAWA INDIA LIMITED" or siteid=="Berger Paints" or siteid=="MK Enterprise" or siteid=="Merlin Projects Ltd" or siteid=="Shree Mangal Trust":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical" or siteid=="Narayan Hospital" or siteid=="Rameswaram Mansion" or siteid=="Veekay Industries" or siteid=="MK MKEE 03":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Sundrex Oil Company Ltd SOCL01" or siteid=="Calcutta Lions Netra Niketan CLSN00001" or siteid=="Future Merchandising FVRL05" or siteid=="Tata Asset Management TTAL01":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Cygnet Industries Ltd KSIL01"or siteid=="JJ Gold Refinery JGLD01" or siteid=="MUMUSO Retail MSLD 01" or siteid=="Sanjay Creation SCIN 01" or siteid=="City Bazaar Metro CIBM01":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
		# print("[rogram is here")
	elif siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16" or siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="Astral Flat Owners ASAF02":
		filename = os.path.join(module_dir,"onlyshift.xlsx")

	elif siteid=="South Point High School SPSC01" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA" or siteid=="TRIPURA ENCLAVE TRPR01" or siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Future Merchandising FVRL12":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Anjali Jewellers ANJL 03" or siteid=="Randstad India M012 01":
		filename = os.path.join(module_dir,"onlyshift.xlsx")

	elif siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="ACC CEMENT ACCL01" or siteid=="Pioneer Packaging PPCI02" or siteid=="EMAMI REALTY LTD EIL201" or siteid=="Chandrani Enterprises CNDR03":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Hiland Sapphire":
		filename = os.path.join(module_dir,"onlyshift1.xlsx")

	elif siteid=="SE Builders And Realtors UKAC01" or siteid=="SE Builders And Realtors UKAC02" or siteid=="Anjali Jewellers ANJL 04" or siteid=="BD Memorial School BDML04" or siteid=="BD Memorial School BDML05":
		filename = os.path.join(module_dir,"onlyshift.xlsx")


	elif siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Fortune Marketing FMTG01" or siteid=="Chandrani Enterprises KOL CNDR01" or siteid=="Chandrani Enterprises KOL CNDR02" or siteid=="Pulse Diagnostics PLSE03" or siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01" or siteid=="Emami Limited GUW EMMI06" or siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63" or siteid=="Wow Momo Foods SIL MOMO01":
		filename = os.path.join(module_dir,"onlyshift.xlsx")
	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06" or siteid=="Spencer Retail SPRL07" or siteid=="Spencer Retail SPRL67":
		filename = os.path.join(module_dir,"onlyshift.xlsx")






	else:
		filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift (1).xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	# ws1 = wb2["Shift A"]
	if siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="YOKOGAWA INDIA LIMITED" or siteid=="Berger Paints" or siteid=="MK Enterprise" or siteid=="Merlin Projects Ltd" or siteid=="Shree Mangal Trust" or siteid=="MK MKEE 03":
		ws2 = wb2["Shift A"]
		if siteid=="Berger Paints" or siteid=="Merlin Projects Ltd" or siteid=="Shree Mangal Trust":
			ws2.title="Night Shift"
	elif siteid=="JJ Gold Refinery JGLD01" or siteid=="MUMUSO Retail MSLD 01" or siteid=="Cygnet Industries Ltd KSIL01" or siteid=="Sanjay Creation SCIN 01" or siteid=="City Bazaar Metro CIBM01" or siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Fortune Marketing FMTG01":
		ws2 = wb2["Shift A"]

	elif siteid=="Tata Asset Management TTAL01" or siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Future Merchandising FVRL12" or siteid=="Randstad India M012 01" or siteid=="Pulse Diagnostics PLSE03" or siteid=="Chandrani Enterprises CNDR03" or siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04":
		ws2 = wb2["Shift A"]


	elif siteid=="Hiland Sapphire" or siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical" or siteid=="Narayan Hospital" or siteid=="Rameswaram Mansion" or siteid=="Veekay Industries":
		ws2 = wb2["Shift A"]
		ws2.title="Night Shift"
	elif siteid=="Sundrex Oil Company Ltd SOCL01" or siteid=="Calcutta Lions Netra Niketan CLSN00001" or siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16":
		ws2 = wb2["Shift A"]
		ws2.title="Night Shift"
	elif siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="Astral Flat Owners ASAF02" or siteid=="Anjali Jewellers ANJL 03":
		ws2 = wb2["Shift A"]
		ws2.title="Night Shift"
	elif siteid=="South Point High School SPSC01" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA" or siteid=="TRIPURA ENCLAVE TRPR01":
		ws2 = wb2["Shift A"]
		ws2.title="Night Shift"
	elif siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01" or siteid=="Emami Limited GUW EMMI06" or siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63" or siteid=="Wow Momo Foods SIL MOMO01":
		ws2 = wb2["Shift A"]

	elif siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="ACC CEMENT ACCL01" or siteid=="Pioneer Packaging PPCI02" or siteid=="EMAMI REALTY LTD EIL201" or siteid=="SE Builders And Realtors UKAC01" or siteid=="SE Builders And Realtors UKAC02" or siteid=="Anjali Jewellers ANJL 04":
		ws2 = wb2["Shift A"]
		ws2.title="Night Shift"

	elif siteid=="Future Merchandising FVRL05" or siteid=="BD Memorial School BDML04" or siteid=="BD Memorial School BDML05" or siteid=="Chandrani Enterprises KOL CNDR01" or siteid=="Chandrani Enterprises KOL CNDR02":
		ws2 = wb2["Shift A"]
		ws2.title="Shift G"
	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06" or siteid=="Spencer Retail SPRL07" or siteid=="Spencer Retail SPRL67":
		ws2 = wb2["Shift A"]






	else:
		ws2 = wb2["Night Shift"]
	
	# ws1['A4'] = siteid
	ws2['A4'] = siteid
	
	# ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
	if siteid=="Teloijan-Tea" or siteid=="South City Retreat SCKL16" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="Astral Flat Owners ASAF02" or siteid=="SE Builders And Realtors UKAC01" or siteid=="SE Builders And Realtors UKAC02" or siteid=="Anjali Jewellers ANJL 04":
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
	elif siteid=="Vikram Solar" or siteid=="Merlin Projects Ltd" or siteid=="Shree Mangal Trust" or siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16":
		ws2['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"

	elif siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="EMAMI REALTY LTD EIL201":
		ws2['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"

	elif siteid=="Future Merchandising FVRL05" or siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"

	elif siteid=="Chandrani Enterprises CNDR03" or siteid=="Emami Limited GUW EMMI06":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 9:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"


	elif siteid=="Pioneer Packaging PPCI02":
		ws2['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"


	elif siteid=="MK Enterprise":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"

	elif siteid=="MK MKEE 03":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:30 Hours"

	elif siteid=="Hiland Sapphire" or siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"

	elif siteid=="Narayan Hospital" or siteid=="TRIPURA ENCLAVE TRPR01" or siteid=="ACC CEMENT ACCL01":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 23:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
	elif siteid=="Rameswaram Mansion":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 05:00 Hours"
	elif siteid=="Veekay Industries":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"

	elif siteid=="Calcutta Lions Netra Niketan CLSN00001":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"

	elif siteid=="Cygnet Industries Ltd KSIL01":
		ws2['A5'] = "General Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"

	elif siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Chandrani Enterprises KOL CNDR02":
		ws2['A5'] = "General Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 9:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
	elif siteid=="Future Merchandising FVRL12" or siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Chandrani Enterprises KOL CNDR01":
		ws2['A5'] = "General Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 9:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
	elif siteid=="Anjali Jewellers ANJL 03":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 09:00 Hours"

	elif siteid=="BD Memorial School BDML04" or siteid=="Wow Momo Foods SIL MOMO01":
		ws2['A5'] = "General Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 7:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	elif siteid=="Pulse Diagnostics PLSE03":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 8:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"

	elif siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 8:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"

	elif siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 7:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 23:00 Hours"
	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 6:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
	elif siteid=="Spencer Retail SPRL07":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 7:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
	elif siteid=="Spencer Retail SPRL67":
		ws2['A5'] = "A Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 6:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 23:00 Hours"



















	elif siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="JJ Gold Refinery JGLD01" or siteid=="Sanjay Creation SCIN 01" or siteid=="Tata Asset Management TTAL01" or siteid=="Randstad India M012 01" or siteid=="BD Memorial School BDML05":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
	elif siteid=="City Bazaar Metro CIBM01":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	elif siteid=="South Point High School SPSC01" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA":
		ws2['A5'] = "Night Shift :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"


	elif siteid=="YOKOGAWA INDIA LIMITED":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 16:00 Hours"
	elif siteid=="MUMUSO Retail MSLD 01" or siteid=="Fortune Marketing FMTG01":
		ws2['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	elif siteid=="Berger Paints":
		ws2['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:30 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:30 Hours"

	else:
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"

	thin = Side (border_style="thin", color="000000")

	if siteid=="CFS PHARMA AND OPTICALS" or siteid=="Ganesh Grains Ltd" or siteid=="YOKOGAWA INDIA LIMITED" or siteid=="MK Enterprise" or siteid=="MK MKEE 03" or siteid=="Cygnet Industries Ltd KSIL01" or siteid=="MUMUSO Retail MSLD 01" or siteid=="JJ Gold Refinery JGLD01" or siteid=="City Bazaar Metro CIBM01" or siteid=="Sanjay Creation SCIN 01":
		rownum1=19


		for ele_ta in timeshiftA:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1

	elif siteid=="Spencer Retail SPRL05" or siteid=="Spencer Retail SPRL06" or siteid=="Spencer Retail SPRL07" or siteid=="Spencer Retail SPRL67":
		rownum1=19


		for ele_ta in timeshiftA:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1


	elif siteid=="Future Merchandising FVRL05" or siteid=="Tata Asset Management TTAL01" or siteid=="Aeropict Logistics Solutions ARLG01" or siteid=="Future Merchandising FVRL12" or siteid=="Randstad India M012 01":
		rownum1=19


		for ele_ta in timeshiftA:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1

	elif siteid=="AGS Retails Ind Ltd AGSL01" or siteid=="Fortune Marketing FMTG01" or siteid=="BD Memorial School BDML04" or siteid=="BD Memorial School BDML05" or siteid=="Chandrani Enterprises KOL CNDR01" or siteid=="Chandrani Enterprises KOL CNDR02" or siteid=="Pulse Diagnostics PLSE03":
		rownum1=19


		for ele_ta in timeshiftA:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1

	elif siteid=="Chandrani Enterprises CNDR03" or siteid=="Silverleaf Tea SRLT01" or siteid=="Ganesh Grains GHGS04" or siteid=="Nemichand Bamalwa Jewellers NMCJ01" or siteid=="Priyanka Agencies EMMI01" or siteid=="Emami Limited GUW EMMI06" or siteid=="Spencer Retail SPRL82" or siteid=="Spencer Retail SPRL63" or siteid=="Wow Momo Foods SIL MOMO01":
		rownum1=19


		for ele_ta in timeshiftA:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1




	elif siteid=="Berger Paints" or siteid=="Merlin Projects Ltd" or siteid=="Shree Mangal Trust" or siteid=="EMAMI FRANK ROSS EMMI13" or siteid=="Astral Flat Owners ASAF02" or siteid=="Anjali Jewellers ANJL 03":
		rownum1=19


		for ele_ta in timeshiftB:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1


	elif siteid=="KISHORILAL SHYAMSUNDER KSSR06" or siteid=="MERLIN RIVER VIEW MRAO01" or siteid=="South City Retreat SCKL16" or siteid=="MERLIN RESIDENCY APPARTMENT MRAA01" or siteid=="SE Builders And Realtors UKAC01" or siteid=="SE Builders And Realtors UKAC02":
		rownum1=19


		for ele_ta in timeshiftB:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1

	elif siteid=="India Auto Mobile" or siteid=="Emami Foundation" or siteid=="ACC CEMENT ACCL01" or siteid=="Pioneer Packaging PPCI02" or siteid=="EMAMI REALTY LTD EIL201" or siteid=="Anjali Jewellers ANJL 04":
		rownum1=19


		for ele_ta in timeshiftB:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1




	elif siteid=="Hiland Sapphire" or siteid=="Orbit Tower" or siteid=="VOLTAMP Electrical" or siteid=="Narayan Hospital" or siteid=="Rameswaram Mansion" or siteid=="Veekay Industries" or siteid=="Sundrex Oil Company Ltd SOCL01" or siteid=="Calcutta Lions Netra Niketan CLSN00001":
		rownum1=19


		for ele_ta in timeshiftB:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1

	elif siteid=="South Point High School SPSC01" or siteid=="Vektra Engineering VKTR01" or siteid=="INTERNATIONAL SCHOOL OF BUSINESS AND MEDIA" or siteid=="TRIPURA ENCLAVE TRPR01":
		rownum1=19


		for ele_ta in timeshiftB:
			ele_ta1=ele_ta.split("-")
			ws2.cell(row=rownum1,column=1).value=ele_ta1[0]
			ws2.cell(row=rownum1,column=2).value=ele_ta1[1]
			ws2.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
			rownum1=rownum1+1


	else:
		pass




	# shiftAcol = 3
	# shiftAi = 0
	# for post in place_list_siteid:
	# 	shiftArow = 19
	# 	for data in isDataShiftA[shiftAi]:
	# 		# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
	# 		# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
	# 		shiftArow += 1
	# 	shiftAi += 1
	# 	shiftAcol += 1




	shiftBcol = 3
	shiftBi = 0
	for post in place_list_siteid:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	col = 3
	for post in range(0, len(place_list_siteid)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteid[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	# ws1['C17'] = 'Post Name'
	ws2['C17'] = 'Post Name'
	# ws1['C17'].font = Font(bold=True)
	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteid)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1

	# ws1['E8'] = place_list_siteid[unAttendedShiftA.index(maxUShiftA)]
	# ws1['E9'] = place_list_siteid[AttendedShiftA.index(maxAShiftA)]
	# ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteid[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteid[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')

	# saving the destination excel file
	if siteid=="Hiland Sapphire":
		# wb2.remove(wb2['Shift A'])

		ws3 = wb2["Timestamp"]

		time1=datetime.now()
		time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		time2_1=time2_1+timedelta(hours=8)
		time2_11=time2_1-timedelta(days=1)

		start_time1=tz.localize(time2_11)
		end_time1=tz.localize(time2_1)

		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=siteid,time__range=(start_time1,end_time1)).values()))

		if len(all_path)!=0:

			all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
			all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

			df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')


			all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
			all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

			df_all=pd.merge(df_all1,all_emp,on='empid_id')

			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			# df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
			df_all=df_all.sort_values("time",ascending=True)
			df_all=df_all.reset_index()

			# print(df_all.columns)

			rownum=2

			for idx,rows in df_all.iterrows():
				ws3.cell(row = rownum, column = 1, value=rows['name'])
				ws3.cell(row = rownum, column = 2, value=rows['placeid_id'])
				ws3.cell(row = rownum, column = 3, value=rows['datetime1'])
				ws3.cell(row = rownum, column = 4, value=rows['empname'])

				ws3.cell(row = rownum, column = 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 2).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 3).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 4).border = Border(top=thin, left=thin, right=thin, bottom=thin)

				rownum=rownum+1



			wb2.save(str(filename1))
		else:
			wb2.save(str(filename1))



	else:

		wb2.save(str(filename1))



	if siteid=="Teloijan-Tea":
		sitenamemail="Poddar House(Teloijan-Tea)"
	else:
		sitenamemail=siteid


	try:
		print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nP.S: This is an auto generated email. Please do not reply on this email.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()


	return HttpResponse("Success")


def qrReportbhu(request):
	siteid = request.GET.get('sitecode',"")
	ret_type=request.GET.get('type',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="Ashwini Hospital":
		notavA=["Casualty","ICCU","Nursing College"]
		notavB=["Casualty","ICCU","Nursing College"]
		notavC=["Lobby Glass Door-2","ICCU"]
		notavG=["Main Gate","Lobby Glass Door-1","Lobby Glass Door-2","Nursing College"]
		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00)]

		shiftC = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]

		shiftG = [(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00)]

	# elif siteid=="AMRI-Bhubaneswar":
	# 	shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00)]

	# 	shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00)]

	# 	shiftC = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]

	# 	shiftG = [(8,00,9,00),(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),
	# 	(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

	elif siteid=="AMRI-Bhubaneswar":
		shiftA = [(6,00,8,00),(8,00,10,00),(10,00,12,00), (12,00,14,00)]

		shiftB = [(14,00,16,00),(16,00,18,00),(18,00,20,00),(20,00,22,00)]

		shiftC = [(22,00,00,00),(00,00,2,00), (2,00,4,00), (4,00,6,00)]

		shiftG = [(8,00,10,00),(10,00,12,00), (12,00,14,00),(14,00,16,00),(16,00,18,00),(18,00,20,00),(20,00,22,00)]


	else:
		notavA=[]
		notavB=[]
		notavC=[]
		notavG=[]

		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00)]

		shiftC = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]

		shiftG = [(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00)]

	# print("Place List--->>>", place_list)
	# place_list_siteid = [i.name+' ('+siteid+')' for i in activeQR]
	
	
	allPlace = []
	uuidListA = []
	uuidListB = []
	uuidListC = []
	uuidListG = []
	place_list_siteidA = []
	place_list_siteidB = []
	place_list_siteidC = []
	place_list_siteidG = []


	if siteid=="Ashwini Hospital" or siteid=="AMRI-Bhubaneswar":

		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ABC")|Q(other="AB"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ABC")|Q(other="AB")|Q(other="AB"),placeid__siteid__name__icontains=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ABC")|Q(other="GC")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="G")|Q(other="GC"),placeid__siteid__name__icontains=siteid).values()))




	uuidListA=df_A["uuid"].tolist()
	df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
	place_List_siteidA=df_A["fullname"].tolist()

	for ele1 in place_List_siteidA:
		ele1=ele1[0:(len(ele1)-len(siteid)-2)]
		place_list_siteidA.append(ele1)



	uuidListB=df_B["uuid"].tolist()
	df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
	place_List_siteidB=df_B["fullname"].tolist()

	for ele2 in place_List_siteidB:
		ele2=ele2[0:(len(ele2)-len(siteid)-2)]
		place_list_siteidB.append(ele2)

	uuidListC=df_C["uuid"].tolist()
	df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
	place_List_siteidC=df_C["fullname"].tolist()

	for ele3 in place_List_siteidC:
		ele3=ele3[0:(len(ele3)-len(siteid)-2)]
		place_list_siteidC.append(ele3)

	uuidListG=df_G["uuid"].tolist()
	df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
	place_List_siteidG=df_G["fullname"].tolist()

	for ele4 in place_List_siteidG:
		ele4=ele4[0:(len(ele4)-len(siteid)-2)]
		place_list_siteidG.append(ele4)





	# print(uuidlistA,place_list_siteidA)
	# print(df_A["name"])





	



	# print("allPlace--->>>", allPlace , len(allPlace))
	# postNo = []
	# for i in range(1, len(place_list_siteid)+1):
	# 	postNo.append('Post No.'+str(i))
	# lenList = [len(i) for i in place_list_siteid]
	# One Day Post Monitoring Report 10 to 10
	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	# shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
	# 			(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
	# 			(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []
	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

	isDataShiftB = []
	AttendedShiftB = []
	unAttendedShiftB = []
	for qrid in uuidListB:
		isQrData = []
		for t in shiftB:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftB.append(isQrData.count(0))
		AttendedShiftB.append(isQrData.count(1))
		isDataShiftB.append(isQrData)


	if siteid=="AMRI-Bhubaneswar":


		isDataShiftC = []
		AttendedShiftC = []
		unAttendedShiftC = []
		for qrid in uuidListC:
			isQrData = []
			for t in shiftC:
				if t == (22,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftC.index(t) > shiftC.index((22,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftC.append(isQrData.count(0))
			AttendedShiftC.append(isQrData.count(1))
			isDataShiftC.append(isQrData)

	else:
		isDataShiftC = []
		AttendedShiftC = []
		unAttendedShiftC = []
		for qrid in uuidListC:
			isQrData = []
			for t in shiftC:
				if t == (23,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftC.index(t) > shiftC.index((23,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftC.append(isQrData.count(0))
			AttendedShiftC.append(isQrData.count(1))
			isDataShiftC.append(isQrData)



	isDataShiftG = []
	AttendedShiftG = []
	unAttendedShiftG = []
	for qrid in uuidListG:
		isQrData = []
		for t in shiftG:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftG.append(isQrData.count(0))
		AttendedShiftG.append(isQrData.count(1))
		isDataShiftG.append(isQrData)





	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	maxAShiftC = max(AttendedShiftC)
	maxUShiftC = max(unAttendedShiftC)
	maxAShiftG = max(AttendedShiftG)
	maxUShiftG = max(unAttendedShiftG)

	if ret_type=="json" and siteid=="AMRI-Bhubaneswar":
		timing_list=["08:00-21:00 hrs","09:00-21:00 hrs","10:00-19:00 hrs","09:00-21:00 hrs","09:00-18:00 hrs","08:00-20:00 hrs","06:00-18:00 hrs","09:00-21:00 hrs",
		"09:00-18:00 hrs","09:00-21:00 hrs","08:00-20:00 hrs","08:00-16:00 hrs","09:00-21:00 hrs","10:00-18:00 hrs","10:00-18:00 hrs"]

		return JsonResponse({"timing_list":timing_list,"AttendedShiftA":AttendedShiftA,"unAttendedShiftA":unAttendedShiftA,"AttendedShiftB":AttendedShiftB,"unAttendedShiftB":unAttendedShiftB,
			"unAttendedShiftC":unAttendedShiftC,"AttendedShiftC":AttendedShiftC,"AttendedShiftG":AttendedShiftG,"unAttendedShiftG":unAttendedShiftG,"maxAShiftA":maxAShiftA,"maxAShiftB":maxAShiftB,
			"maxAShiftC":maxAShiftC,"maxAShiftG":maxAShiftG,"maxUShiftA":maxUShiftA,"maxUShiftB":maxUShiftB,"maxUShiftC":maxUShiftC,"maxUShiftG":maxUShiftG,"isDataShiftA":isDataShiftA,
			"isDataShiftB":isDataShiftB,"isDataShiftC":isDataShiftC,"isDataShiftG":isDataShiftG,"place_list_siteidA":place_list_siteidA,"place_list_siteidB":place_list_siteidB,
			"place_list_siteidC":place_list_siteidC,"place_list_siteidG":place_list_siteidG})
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)
	if siteid=="AMRI-Bhubaneswar":
		filename = os.path.join(module_dir,"Post Monitoring Report Format_abhu.xlsx")

	else:
		filename = os.path.join(module_dir,"Post Monitoring Report Format_all.xlsx")
	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Shift A"]
	ws2 = wb2["Shift B"]
	ws3 = wb2["Shift C"]
	ws4 = wb2["Shift G"]
	ws1['A4'] = siteid
	ws2['A4'] = siteid
	ws3['A4'] = siteid
	ws4['A4'] = siteid

	if siteid=="Ashwini Hospital":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"
	elif siteid=="AMRI-Bhubaneswar":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"



	thin = Side (border_style="thin", color="000000")

	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	shiftCcol = 3
	shiftCi = 0
	for postC in place_list_siteidC:
		shiftCrow = 19
		for data in isDataShiftC[shiftCi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftCrow += 1
		shiftCi += 1
		shiftCcol += 1


	shiftGcol = 3
	shiftGi = 0
	for postG in place_list_siteidG:
		shiftGrow = 19
		for data in isDataShiftG[shiftGi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftGrow += 1
		shiftGi += 1
		shiftGcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteidC[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws3['C17'] = 'Post Name'

	ws3['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws4.cell(row = 18, column = col, value=place_list_siteidG[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws4.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws4.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws4.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws4.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	if siteid=="AMRI-Bhubaneswar":
		ws4.merge_cells(start_row = 16, end_row=16, start_column=3, end_column = col-1)

		ws4['C16'] = 'Post Availability Timing & Post Name'

		ws4['C16'].font = Font(bold=True)

		# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		ws4.cell(row = 16, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 16, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

		timing_list=["08:00-21:00 hrs","09:00-21:00 hrs","10:00-19:00 hrs","09:00-21:00 hrs","09:00-18:00 hrs","08:00-20:00 hrs","06:00-18:00 hrs","09:00-21:00 hrs",
		"09:00-18:00 hrs","09:00-21:00 hrs","08:00-20:00 hrs","08:00-16:00 hrs","09:00-21:00 hrs","10:00-18:00 hrs","10:00-18:00 hrs"]

		rn=17
		cn=3
		k=0

		while k<len(timing_list):
			ws4.cell(row =rn, column = cn).value = timing_list[k]
			k=k+1
			cn+=1



	else:
		ws4.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

		ws4['C17'] = 'Post Name'

		ws4['C17'].font = Font(bold=True)

		# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		ws4.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws4.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1



	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
	ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['B18'].alignment = Alignment(horizontal='center', vertical='center')

	# saving the destination excel file
	wb2.save(str(filename1))

	sitenamemail=siteid

	try:
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily Post Monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		# mail.attach_file(filename2)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	return HttpResponse("success")

def qrReportEmami(request):
	siteid = request.GET.get('sitecode')
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]
	place_list_siteid = []

	allPlace = []
	uuidList = []
	# lenSite = len(siteid)
	# # place_list_siteid_1 = [i.replace(f'({siteid})', '') for i in place_list]
	# # place_list_siteid_short = []
	# for place in place_list:	
	# 	qrList = uuidTbl.objects.filter(placeid_id=place)
	# 	for placeuid in qrList:
	# 		place_list_siteid.append("( "+placeuid.name+" ) "+place)
	# 		allPlace.append(placeuid.name)
	# 		uuidList.append(placeuid.uuid)
	tz=timezone.get_current_timezone()
	
	allPost = []
	allPatrol = []
	AllPost = []
	AllPatrol = []
	for place in place_list:
		qrList = uuidTbl.objects.filter(placeid_id=place)
		for placeuid in qrList:
			place_list_siteid.append("( "+placeuid.name+" ) "+place)
			allPlace.append(placeuid.name)
			uuidList.append((placeuid.uuid, placeuid.uuidtype, placeuid.name))
			if placeuid.uuidtype == 'QR':
				allPost.append(placeuid.name)
				AllPost.append(placeuid.name+" ["+place+ "]")
			else:
				allPatrol.append(placeuid.name)
				AllPatrol.append(placeuid.name+" ["+place+ "]")
	# for place in place_list:
	# 	qrList = uuidTbl.objects.filter(placeid_id=place)
	# 	for placeuid in qrList:
	# 		place_list_siteid.append("( "+placeuid.name+" ) "+place)
	# 		allPlace.append(placeuid.name)
	# 		uuidList.append((placeuid.uuid, placeuid.uuidtype, placeuid.name))
	# 		if placeuid.uuidtype == 'QR':
	# 			allPost.append(placeuid.name)
	# 		else:
	# 			allPatrol.append(placeuid.name)

	lenList = [len(i) for i in place_list_siteid]
	# One Day Post Monitoring Report 10 to 10
	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	dt_1=yest.strftime("%d-%b-%Y")
	yestStr = yest.strftime('%Y %d %m').split()
	
	AmingaonShiftANA = ['Back Side DG Area', 'Store Area']
	AmingaonShiftBNA = ['Lady Security Area', 'Lady Security Area', 'Lady Security Area', 'Lady Security Area', 'Store Area']
	NA_shiftA_amingaon = ['Back Side DG Area [Area-Outside(Emami-Amingaon)]', 'Store Area [Store-Gr Flr(Emami-Amingaon)]']
	NA_shiftB_amingaon = ['Lady Security Area [Production-1st Floor(Emami-Amingaon)]', 'Lady Security Area [Production-1st Flr West(Emami-Amingaon)]', 'Lady Security Area [Production-Gr Flr West(Emami-Amingaon)]', 'Lady Security Area [Production-Ground Floor(Emami-Amingaon)]', 'Store Area [Store-Gr Flr(Emami-Amingaon)]']

	AbhoypurShiftANA=['Factory Front Side', 'Back Side Utility Area']
	NA_shiftA_abhoypur=['Factory Front Side [Area-Outside(Emami-Abhoypur)]', 'Back Side Utility Area [Area-Outside(Emami-Abhoypur)]']
	
	post_shiftA_ap = [(6,00,6,30), (6,30,7,00), (7,00,7,30), (7,30,8,00), (8,00,8,30),
				(8,30,9,00), (9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00),
				(11,00,11,30), (11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30),
				(13,30,14,00), (14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00),
				(16,00,16,30), (16,30,17,00), (17,00,17,30), (17,30,18,00)]

	post_shiftB_ap = [(18,00,18,30), (18,30,19,00), (19,00,19,30), (19,30,20,00),
				(20,00,20,30), (20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),
				(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),
				(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00)]

	post_shiftA_ag = [(6,00,6,30), (6,30,7,00),(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00),
				(9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00), (11,00,11,30),
				(11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00),
				(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
				(16,30,17,00), (17,00,17,30), (17,30,18,00)]

	post_shiftB_ag = [(18,00,18,30), (18,30,19,00),(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00),
				(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
				(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), 
				(2,00,2,30), (2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00),
				(5,00,5,30), (5,30,6,00)]
	
	
	patrol_abhoypur_shiftA = [(6,00,8,00), (8,00,10,00), (10,00,12,00), (12,00,14,00), (14,00,16,00), (16,00,18,00)]
	patrol_abhoypur_shiftB = [(18,00,20,00), (20,00,22,00), (22,00,00,00), (00,00,2,00), (2,00,4,00), (4,00,6,00)]
	
	patrol_amingaon_shiftA = [(6,00,8,00), (8,00,10,00), (10,00,12,00), (12,00,14,00), (14,00,16,00), (16,00,18,00)]
	patrol_amingaon_shiftB = [(18,00,20,00), (20,00,22,00), (22,00,00,00), (00,00,2,00), (2,00,4,00), (4,00,6,00)]

	# patrol_pacharia_shiftA = [(6,00,9,00), (9,00,12,00), (12,00,15,00), (15,00,18,00)]
	patrol_pacharia_shiftA=[(8,00,11,00), (11,00,14,00), (14,00,17,00), (17,00,20,00)] 
	# patrol_pacharia_shiftB = [(18,00,21,00), (21,00,00,00), (00,00,3,00), (3,00,6,00)]
	patrol_pacharia_shiftB=[(20,00,23,00), (23,00,2,00), (2,00,5,00), (5,00,8,00)]

	isPostDataShiftA = []
	isPatrolDataShiftA = []
	PostAttendedShiftA = []
	PostunAttendedShiftA = []
	PatrolAttendedShiftA = []
	PatrolunAttendedShiftA = []
	PatrolCountShiftA = []

	for qrid in uuidList:
		isPostQrData = []
		isPatrolQrData = []
		PatrolCount = []
		if siteid == "Emami-Abhoypur":
			if qrid[1] == 'QR':
				if qrid[-1] not in AbhoypurShiftANA:
					for t in post_shiftA_ap:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					PostunAttendedShiftA.append(isPostQrData.count(0))
					PostAttendedShiftA.append(isPostQrData.count(1))
					isPostDataShiftA.append(isPostQrData)
			if qrid[1] == 'QR PATROL':
				for t in patrol_abhoypur_shiftA:
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
					if data:
						isPatrolQrData.append(1)
						PatrolCount.append(len(data))
					else:
						isPatrolQrData.append(0)
						PatrolCount.append(0)
				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
				isPatrolDataShiftA.append(isPatrolQrData)
				PatrolCountShiftA.append(PatrolCount)

		if siteid == "Emami-Amingaon":
			if qrid[1] == 'QR':
				if qrid[2] not in AmingaonShiftANA:
					for t in post_shiftA_ag:					
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					PostunAttendedShiftA.append(isPostQrData.count(0))
					PostAttendedShiftA.append(isPostQrData.count(1))
					isPostDataShiftA.append(isPostQrData)
			if qrid[1] == 'QR PATROL':
				for t in patrol_amingaon_shiftA:
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
					if data:
						PatrolCount.append(len(data))
						isPatrolQrData.append(1)
					else:
						PatrolCount.append(0)
						isPatrolQrData.append(0)
				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
				isPatrolDataShiftA.append(isPatrolQrData)
			PatrolCountShiftA.append(PatrolCount)

		if siteid == "Emami-Pacharia":
			if qrid[1] == 'QR':
				for t in post_shiftA_ap:
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
					if data:
						isPostQrData.append(1)
					else:
						isPostQrData.append(0)
				PostunAttendedShiftA.append(isPostQrData.count(0))
				PostAttendedShiftA.append(isPostQrData.count(1))
				isPostDataShiftA.append(isPostQrData)
			if qrid[1] == 'QR PATROL':
				for t in patrol_pacharia_shiftA:
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
					if data:
						PatrolCount.append(len(data))
						isPatrolQrData.append(1)
					else:
						PatrolCount.append(0)
						isPatrolQrData.append(0)
				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
				isPatrolDataShiftA.append(isPatrolQrData)
				PatrolCountShiftA.append(PatrolCount)

	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

	isPostDataShiftB = []
	isPatrolDataShiftB = []
	PostAttendedShiftB = []
	PostunAttendedShiftB = []
	PatrolAttendedShiftB = []
	PatrolunAttendedShiftB = []
	PatrolCountShiftB = []

	for qrid in uuidList:
		isPostQrData = []
		isPatrolQrData = []
		patrolCount = []

		if siteid == "Emami-Abhoypur":
			if qrid[1] == 'QR':
				for t in post_shiftB_ap:
					if t == (23,30,00,00):
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					elif post_shiftB_ap.index(t) > post_shiftB_ap.index((23,30,00,00)):
						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					else:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)

				PostunAttendedShiftB.append(isPostQrData.count(0))
				PostAttendedShiftB.append(isPostQrData.count(1))
				isPostDataShiftB.append(isPostQrData)
			if qrid[1] == 'QR PATROL':
				for t in patrol_abhoypur_shiftB:
					if t == (22,00,00,00):
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					elif patrol_abhoypur_shiftB.index(t) > patrol_abhoypur_shiftB.index((22,00,00,00)):
						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					else:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)

				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
				isPatrolDataShiftB.append(isPatrolQrData)

		if siteid == "Emami-Pacharia":
			if qrid[1] == 'QR':
				for t in post_shiftB_ap:
					if t == (23,30,00,00):
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					elif post_shiftB_ap.index(t) > post_shiftB_ap.index((23,30,00,00)):
						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)
					else:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPostQrData.append(1)
						else:
							isPostQrData.append(0)

				PostunAttendedShiftB.append(isPostQrData.count(0))
				PostAttendedShiftB.append(isPostQrData.count(1))
				isPostDataShiftB.append(isPostQrData)
			
			if qrid[1] == 'QR PATROL':
				for t in patrol_pacharia_shiftB:
					# if t == (21,00,00,00):
					if t == (23,00,2,00):
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					# elif patrol_pacharia_shiftB.index(t) > patrol_pacharia_shiftB.index((21,00,00,00)):
					elif patrol_pacharia_shiftB.index(t) > patrol_pacharia_shiftB.index((23,00,2,00)):
						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					else:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)


				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
				isPatrolDataShiftB.append(isPatrolQrData)

		if siteid == "Emami-Amingaon":
			if qrid[1] == 'QR':
				if qrid[2] not in AmingaonShiftBNA:
					for t in post_shiftB_ag:
						if t == (23,30,00,00):
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isPostQrData.append(1)
							else:
								isPostQrData.append(0)
						elif post_shiftB_ag.index(t) > post_shiftB_ag.index((23,30,00,00)):
							start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isPostQrData.append(1)
							else:
								isPostQrData.append(0)
						else:
							start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
							end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
							start=tz.localize(start)
							end=tz.localize(end)
							data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
							if data:
								isPostQrData.append(1)
							else:
								isPostQrData.append(0)

					PostunAttendedShiftB.append(isPostQrData.count(0))
					PostAttendedShiftB.append(isPostQrData.count(1))
					isPostDataShiftB.append(isPostQrData)

			if qrid[1] == 'QR PATROL':
				for t in patrol_amingaon_shiftB:
					if t == (22,00,0,00):
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					elif patrol_amingaon_shiftB.index(t) > patrol_amingaon_shiftB.index((22,00,0,00)):
						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)
					else:
						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
						start=tz.localize(start)
						end=tz.localize(end)
						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
						if data:
							isPatrolQrData.append(1)
							patrolCount.append(len(data))
						else:
							isPatrolQrData.append(0)
							patrolCount.append(0)

				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
				isPatrolDataShiftB.append(isPatrolQrData)

	PostmaxAShiftA = max(PostAttendedShiftA)
	PostmaxUShiftA = max(PostunAttendedShiftA)
	PatrolmaxAShiftA = max(PatrolAttendedShiftA)
	PatrolmaxUShiftA = max(PatrolunAttendedShiftA)

	PostmaxAShiftB = max(PostAttendedShiftB)
	PostmaxUShiftB = max(PostunAttendedShiftB)
	PatrolmaxAShiftB = max(PatrolAttendedShiftB)
	PatrolmaxUShiftB = max(PatrolunAttendedShiftB)
	# print()
	# print("-------------------------------------------------")
	# print(place_list_siteid)
	# print("-----------------All Post---------------------")
	# print(allPost)
	# print("-----------------------All Patrol--------------")
	# print(allPatrol)
	maxLen = int(max(lenList))
	# BASE_DIR = Path(__file__).resolve().parent.parent
	module_dir = os.path.dirname(__file__)
	Postfile = ''
	Patrolfile = ''

	# filename = os.path.join(module_dir,"Post Monitoring Report Format_abhu.xlsx")


	if siteid == "Emami-Abhoypur":
		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Abhoypur.xlsx'))
		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Abhoypur.xlsx'))
	if siteid == "Emami-Amingaon":
		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Amingaon.xlsx'))
		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Amingaon.xlsx'))
	if siteid == "Emami-Pacharia":
		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Pacharia.xlsx'))
		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Pacharia.xlsx'))
	
	wb1Post = xl.load_workbook(Postfile)
	wb1Patrol = xl.load_workbook(Patrolfile)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
	filename2 =siteid+" Patrol Monitoring Report.xlsx"
	wb2Post = xl.load_workbook(Postfile)
	wb2Patrol = xl.load_workbook(Patrolfile)

	wsPostDayShift = wb2Post["Day Shift"]
	wsPostNightShift = wb2Post["Night Shift"]
	wsPatrolDayShift = wb2Patrol["Day Shift"]
	wsPatrolNightShift = wb2Patrol["Night Shift"]

	wsPostDayShift['A4'] = siteid
	wsPostNightShift['A4'] = siteid
	wsPatrolDayShift['A4'] = siteid
	wsPatrolNightShift['A4'] = siteid
	
	if siteid == "Emami-Abhoypur":
		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
	if siteid == "Emami-Amingaon":
		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
	if siteid == "Emami-Pacharia":
		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
	
	finalPost = []
	finalPostNight = []
	finalPostDay = []
	# if siteid == 'Emami-Abhoypur':
	# 	for i in allPost:
	# 		if i not in ('Factory Front Side', 'Back Side Utility Area'):
	# 			finalPost.append(i)
	# elif siteid == 'Emami-Amingaon':
	# 	for i in allPost:
	# 		if i not in ('Back Side DG Area','Loading Point'):
	# 			finalPost.append(i)
	# else:
	# 	for i in allPost:
	# 		finalPost.append(i)

	if siteid == 'Emami-Abhoypur':
		for i in AllPost:
			if i not in NA_shiftA_abhoypur:
			# if i not in ('Factory Front Side', 'Back Side Utility Area'):
			# if 'Factory Front Side' or 'Back Side Utility Area' not in i:
				finalPostDay.append(i)
			finalPostNight.append(i)
	elif siteid == 'Emami-Amingaon':
		for i in AllPost:
			# if i not in ('Back Side DG Area','Loading Point'):
			if i not in NA_shiftA_amingaon:
			# if 'Back Side DG Area' or 'Loading Point' not in i:
				finalPost.append(i)
				finalPostDay.append(i)
			if i not in NA_shiftB_amingaon:
				finalPostNight.append(i)
	else:
		for i in AllPost:
			finalPost.append(i)
			finalPostDay.append(i)
			finalPostNight.append(i)

	# print("---------------------------------------------------------START-----------------------------------------------------------")
	# print("-----------------------------------------All Post------------------------------\n", AllPost)
	# print("\n--------------------------------------Final Post Day-----------------------------------\n", finalPostDay)
	# print("-----------------------------------------------------------END----------------------------------------------------------")
	# # finalPostDay[PostAttendedShiftA.index(PostmaxAShiftA)]
	# print("PostmaxAShiftA--->>>", PostmaxAShiftA)
	# print("PostAttendedShiftA--->>>", PostAttendedShiftA)

	thin = Side(border_style="thin", color="000000")

	# --------------------------Data fill Start Post Monitoring------------------------
	if siteid == 'Emami-Abhoypur' or siteid == 'Emami-Amingaon':
		shiftAcolDay = 3
		shiftAiday = 0
		for post in finalPostDay:
			shiftArowDPost = 19
			for data in isPostDataShiftA[shiftAiday]:
				# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
				# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				if data:
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = '64c620', fill_type="solid")
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				else:
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				shiftArowDPost += 1
			shiftAiday += 1
			shiftAcolDay += 1
		shiftAcol = 3
		shiftAi = 0
		for post in finalPostNight:
			shiftArowNPost = 19
			for data in isPostDataShiftB[shiftAi]:
				# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
				# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				if data:
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				else:
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				shiftArowNPost += 1
			shiftAi += 1
			shiftAcol += 1
	else:
		shiftAcolDay = 3
		shiftAiday = 0
		for post in finalPostDay:
			shiftArowDPost = 19
			for data in isPostDataShiftA[shiftAiday]:
				# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
				# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				if data:
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = '64c620', fill_type="solid")
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				else:
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
					wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				shiftArowDPost += 1
			shiftAiday += 1
			shiftAcolDay += 1
		shiftAcol = 3
		shiftAi = 0
		for post in finalPostNight:
			shiftArowNPost = 19
			for data in isPostDataShiftB[shiftAi]:
				# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
				# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
				if data:
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				else:
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
					wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				shiftArowNPost += 1
			shiftAi += 1
			shiftAcol += 1

	# -----------------------END------------------------

	# --------------------------Data fill Start Patrol Monitoring------------------------
	shiftBcol = 3
	shiftBi = 0
	for post in allPatrol:
		shiftBrowDPatrol = 19
		shiftBroNPatrol = 19
		for data in isPatrolDataShiftA[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrowDPatrol += 1
		for data in isPatrolDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBroNPatrol += 1

		shiftBi += 1
		shiftBcol += 1
	# -----------------------END------------------------
	
	# --------------------------Post name fill Start Post Monitoring------------------------
	if siteid == 'Emami-Abhoypur' or siteid == 'Emami-Amingaon':
		col = 3
		for post in range(0, len(finalPostNight)):
			wsPostNightShift.cell(row = 18, column = col, value=finalPostNight[post])
			wsPostNightShift.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			wsPostNightShift.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			wsPostNightShift.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
			wsPostNightShift.column_dimensions[get_column_letter(col)].width = 25
			wsPostNightShift.row_dimensions[18].height = 70
			col += 1
		
		colDay = 3
		for post in range(0, len(finalPostDay)):
			wsPostDayShift.cell(row = 18, column = colDay, value=finalPostDay[post])
			wsPostDayShift.cell(row = 18, column = colDay).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			wsPostDayShift.cell(row = 18, column = colDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			wsPostDayShift.cell(row = 18, column = colDay).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

			wsPostDayShift.column_dimensions[get_column_letter(colDay)].width = 25
			wsPostDayShift.row_dimensions[18].height = 70
			colDay += 1
	else:
		col = 3
		for post in range(0, len(finalPostNight)):
			wsPostNightShift.cell(row = 18, column = col, value=finalPostNight[post])
			wsPostNightShift.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			wsPostNightShift.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			wsPostNightShift.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
			wsPostNightShift.column_dimensions[get_column_letter(col)].width = 25
			wsPostNightShift.row_dimensions[18].height = 70
			col += 1
		
		colDay = 3
		for post in range(0, len(finalPostDay)):
			wsPostDayShift.cell(row = 18, column = colDay, value=finalPostDay[post])
			wsPostDayShift.cell(row = 18, column = colDay).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			wsPostDayShift.cell(row = 18, column = colDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			wsPostDayShift.cell(row = 18, column = colDay).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

			wsPostDayShift.column_dimensions[get_column_letter(colDay)].width = 25
			wsPostDayShift.row_dimensions[18].height = 70
			colDay += 1

	# -----------------------END------------------------
	
	# --------------------------Post name fill Start Patrol Monitoring------------------------
	Patrolcol = 3
	for post in range(0, len(allPatrol)):
		wsPatrolDayShift.cell(row = 18, column = Patrolcol, value=AllPatrol[post])
		wsPatrolNightShift.cell(row = 18, column = Patrolcol, value=AllPatrol[post])
		
		wsPatrolDayShift.cell(row = 18, column = Patrolcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		wsPatrolNightShift.cell(row = 18, column = Patrolcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		wsPatrolDayShift.cell(row = 18, column = Patrolcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPatrolNightShift.cell(row = 18, column = Patrolcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		wsPatrolDayShift.cell(row = 18, column = Patrolcol).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		wsPatrolNightShift.cell(row = 18, column = Patrolcol).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		wsPatrolDayShift.column_dimensions[get_column_letter(Patrolcol)].width = 25
		wsPatrolNightShift.column_dimensions[get_column_letter(Patrolcol)].width = 25
		wsPatrolDayShift.row_dimensions[18].height = 70
		wsPatrolNightShift.row_dimensions[18].height = 70
		Patrolcol += 1
	# -----------------------END------------------------

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	wsPostDayShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = colDay-1)
	wsPostNightShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	
	wsPatrolDayShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = Patrolcol-1)
	wsPatrolNightShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = Patrolcol-1)
	
	# ws1['C17'] = 'Post Name'
	wsPostDayShift['C17'] = 'Post Name'
	wsPostNightShift['C17'] = 'Post Name'
	wsPatrolDayShift['C17'] = 'Post Name'
	wsPatrolNightShift['C17'] = 'Post Name'
	# ws1['C17'].font = Font(bold=True)
	wsPostDayShift['C17'].font = Font(bold=True)
	wsPostNightShift['C17'].font = Font(bold=True)
	wsPatrolDayShift['C17'].font = Font(bold=True)
	wsPatrolNightShift['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	wsPostDayShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	wsPostNightShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	wsPatrolDayShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	wsPatrolNightShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	wsPostDayShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	wsPostNightShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	wsPatrolDayShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	wsPatrolNightShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercolPost = 3
	for post in range(0, len(allPost)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPostNightShift.cell(row = 17, column = bordercolPost).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercolPost += 1
	
	bordercolPostDay = 3
	for post in range(0, len(finalPost)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPostDayShift.cell(row = 17, column = bordercolPostDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercolPostDay += 1

	bordercolPatrol = 3
	for post in range(0, len(allPatrol)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPatrolDayShift.cell(row = 17, column = bordercolPatrol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPatrolNightShift.cell(row = 17, column = bordercolPatrol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercolPatrol += 1

	# ws1['E8'] = place_list_siteid[unAttendedShiftA.index(maxUShiftA)]
	# ws1['E9'] = place_list_siteid[AttendedShiftA.index(maxAShiftA)]
	# ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True
	wsPostDayShift['E8'] = finalPostDay[PostunAttendedShiftA.index(PostmaxUShiftA)]
	wsPostDayShift['E9'] = finalPostDay[PostAttendedShiftA.index(PostmaxAShiftA)]
	wsPostDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPostNightShift['E8'] = finalPostNight[PostunAttendedShiftB.index(PostmaxUShiftB)]
	wsPostNightShift['E9'] = finalPostNight[PostAttendedShiftB.index(PostmaxAShiftB)]
	wsPostNightShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostNightShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolDayShift['E8'] = AllPatrol[PatrolunAttendedShiftA.index(PatrolmaxUShiftA)]
	wsPatrolDayShift['E9'] = AllPatrol[PatrolAttendedShiftA.index(PatrolmaxAShiftA)]
	wsPatrolDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPatrolDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolNightShift['E8'] = AllPatrol[PatrolunAttendedShiftB.index(PatrolmaxUShiftB)]
	wsPatrolNightShift['E9'] = AllPatrol[PatrolAttendedShiftB.index(PatrolmaxAShiftB)]
	wsPatrolNightShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPatrolNightShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostDayShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostDayShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

	wsPostNightShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostNightShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolDayShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
	wsPatrolDayShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolNightShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
	wsPatrolNightShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

	# saving the destination excel file
	wb2Post.save(str(filename1))
	wb2Patrol.save(str(filename2))

	sitenamemail=siteid

	try:
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Remote Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily Post & Patrol Monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.attach_file(filename2)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	

	return HttpResponse("Success")


# def qrReportEmami(request):
# 	siteid = request.GET.get('sitecode')
# 	activeQR = placeTbl.objects.filter(siteid_id=siteid)
# 	place_list = [i.name for i in activeQR]
# 	place_list_siteid = []

# 	allPlace = []
# 	uuidList = []
# 	# lenSite = len(siteid)
# 	# # place_list_siteid_1 = [i.replace(f'({siteid})', '') for i in place_list]
# 	# # place_list_siteid_short = []
# 	# for place in place_list:	
# 	# 	qrList = uuidTbl.objects.filter(placeid_id=place)
# 	# 	for placeuid in qrList:
# 	# 		place_list_siteid.append("( "+placeuid.name+" ) "+place)
# 	# 		allPlace.append(placeuid.name)
# 	# 		uuidList.append(placeuid.uuid)
# 	tz=timezone.get_current_timezone()
	
# 	allPost = []
# 	allPatrol = []
# 	AllPost = []
# 	AllPatrol = []
# 	for place in place_list:
# 		qrList = uuidTbl.objects.filter(placeid_id=place)
# 		for placeuid in qrList:
# 			place_list_siteid.append("( "+placeuid.name+" ) "+place)
# 			allPlace.append(placeuid.name)
# 			uuidList.append((placeuid.uuid, placeuid.uuidtype, placeuid.name))
# 			if placeuid.uuidtype == 'QR':
# 				allPost.append(placeuid.name)
# 				AllPost.append(placeuid.name+" ["+place+ "]")
# 			else:
# 				allPatrol.append(placeuid.name)
# 				AllPatrol.append(placeuid.name+" ["+place+ "]")
# 	# for place in place_list:
# 	# 	qrList = uuidTbl.objects.filter(placeid_id=place)
# 	# 	for placeuid in qrList:
# 	# 		place_list_siteid.append("( "+placeuid.name+" ) "+place)
# 	# 		allPlace.append(placeuid.name)
# 	# 		uuidList.append((placeuid.uuid, placeuid.uuidtype, placeuid.name))
# 	# 		if placeuid.uuidtype == 'QR':
# 	# 			allPost.append(placeuid.name)
# 	# 		else:
# 	# 			allPatrol.append(placeuid.name)

# 	lenList = [len(i) for i in place_list_siteid]
# 	# One Day Post Monitoring Report 10 to 10
# 	todayStr = date.today().strftime('%Y %d %m').split()
# 	today = date.today()
# 	yest = today - timedelta(days = 1)
# 	dt_1=yest.strftime("%d-%b-%Y")
# 	yestStr = yest.strftime('%Y %d %m').split()
	
# 	NA_shiftA_amingaon = ['Back Side DG Area', 'Loading Point']
# 	NA_shiftA_abhoypur = ['Factory Front side', 'Back side Utility area']

# 	post_shiftA_ap = [(6,00,6,30), (6,30,7,00), (7,00,7,30), (7,30,8,00), (8,00,8,30),
# 				(8,30,9,00), (9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00),
# 				(11,00,11,30), (11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30),
# 				(13,30,14,00), (14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00),
# 				(16,00,16,30), (16,30,17,00), (17,00,17,30), (17,30,18,00)]

# 	post_shiftB_ap = [(18,00,18,30), (18,30,19,00), (19,00,19,30), (19,30,20,00),
# 				(20,00,20,30), (20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),
# 				(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),
# 				(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
# 				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00)]

# 	post_shiftA_ag = [(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00),
# 				(9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00), (11,00,11,30),
# 				(11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00),
# 				(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
# 				(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00)]

# 	post_shiftB_ag = [(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00),
# 				(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
# 				(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), 
# 				(2,00,2,30), (2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00),
# 				(5,00,5,30), (5,30,6,00), (6,00,6,30), (6,30,7,00)]
	
	
# 	patrol_abhoypur_shiftA = [(6,00,8,00), (8,00,10,00), (10,00,12,00), (12,00,14,00), (14,00,16,00), (16,00,18,00)]
# 	patrol_abhoypur_shiftB = [(18,00,20,00), (20,00,22,00), (22,00,00,00), (00,00,2,00), (2,00,4,00), (4,00,6,00)]
	
# 	patrol_amingaon_shiftA = [(7,00,9,00), (9,00,11,00), (11,00,13,00), (13,00,15,00), (15,00,17,00), (17,00,19,00)]
# 	patrol_amingaon_shiftB = [(19,00,21,00), (21,00,23,00), (23,00,1,00), (1,00,3,00), (3,00,5,00), (5,00,7,00)]

# 	# patrol_pacharia_shiftA = [(6,00,9,00), (9,00,12,00), (12,00,15,00), (15,00,18,00)]
# 	patrol_pacharia_shiftA=[(8,00,11,00), (11,00,14,00), (14,00,17,00), (17,00,20,00)] 
# 	# patrol_pacharia_shiftB = [(18,00,21,00), (21,00,00,00), (00,00,3,00), (3,00,6,00)]
# 	patrol_pacharia_shiftB=[(20,00,23,00), (23,00,2,00), (2,00,5,00), (5,00,8,00)]

# 	isPostDataShiftA = []
# 	isPatrolDataShiftA = []
# 	PostAttendedShiftA = []
# 	PostunAttendedShiftA = []
# 	PatrolAttendedShiftA = []
# 	PatrolunAttendedShiftA = []
# 	PatrolCountShiftA = []

# 	for qrid in uuidList:
# 		isPostQrData = []
# 		isPatrolQrData = []
# 		PatrolCount = []
# 		if siteid == "Emami-Abhoypur":
# 			print(qrid[-1])
# 			if qrid[1] == 'QR':
# 				if qrid[-1] not in ('Factory Front Side', 'Back Side Utility Area'):
# 					for t in post_shiftA_ap:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					PostunAttendedShiftA.append(isPostQrData.count(0))
# 					PostAttendedShiftA.append(isPostQrData.count(1))
# 					isPostDataShiftA.append(isPostQrData)
# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_abhoypur_shiftA:
# 					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 					start=tz.localize(start)
# 					end=tz.localize(end)
# 					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 					if data:
# 						isPatrolQrData.append(1)
# 						PatrolCount.append(len(data))
# 					else:
# 						isPatrolQrData.append(0)
# 						PatrolCount.append(0)
# 				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftA.append(isPatrolQrData)
# 				PatrolCountShiftA.append(PatrolCount)

# 		if siteid == "Emami-Amingaon":
# 			if qrid[1] == 'QR':
# 				if qrid[2] not in ('Back Side DG Area','Loading Point'):
# 					for t in post_shiftA_ag:					
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					PostunAttendedShiftA.append(isPostQrData.count(0))
# 					PostAttendedShiftA.append(isPostQrData.count(1))
# 					isPostDataShiftA.append(isPostQrData)
# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_amingaon_shiftA:
# 					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 					start=tz.localize(start)
# 					end=tz.localize(end)
# 					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 					if data:
# 						PatrolCount.append(len(data))
# 						isPatrolQrData.append(1)
# 					else:
# 						PatrolCount.append(0)
# 						isPatrolQrData.append(0)
# 				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftA.append(isPatrolQrData)
# 			PatrolCountShiftA.append(PatrolCount)

# 		if siteid == "Emami-Pacharia":
# 			if qrid[1] == 'QR':
# 				for t in post_shiftA_ap:
# 					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 					start=tz.localize(start)
# 					end=tz.localize(end)
# 					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 					if data:
# 						isPostQrData.append(1)
# 					else:
# 						isPostQrData.append(0)
# 				PostunAttendedShiftA.append(isPostQrData.count(0))
# 				PostAttendedShiftA.append(isPostQrData.count(1))
# 				isPostDataShiftA.append(isPostQrData)
# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_pacharia_shiftA:
# 					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 					start=tz.localize(start)
# 					end=tz.localize(end)
# 					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid[0])
# 					if data:
# 						PatrolCount.append(len(data))
# 						isPatrolQrData.append(1)
# 					else:
# 						PatrolCount.append(0)
# 						isPatrolQrData.append(0)
# 				PatrolunAttendedShiftA.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftA.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftA.append(isPatrolQrData)
# 				PatrolCountShiftA.append(PatrolCount)

# 	# print("uuidList--->>>", uuidList, len(uuidList))
# 	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

# 	isPostDataShiftB = []
# 	isPatrolDataShiftB = []
# 	PostAttendedShiftB = []
# 	PostunAttendedShiftB = []
# 	PatrolAttendedShiftB = []
# 	PatrolunAttendedShiftB = []
# 	PatrolCountShiftB = []

# 	for qrid in uuidList:
# 		isPostQrData = []
# 		isPatrolQrData = []
# 		patrolCount = []

# 		if siteid == "Emami-Abhoypur":
# 			if qrid[1] == 'QR':
# 				for t in post_shiftB_ap:
# 					if t == (23,30,00,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					elif post_shiftB_ap.index(t) > post_shiftB_ap.index((23,30,00,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)

# 				PostunAttendedShiftB.append(isPostQrData.count(0))
# 				PostAttendedShiftB.append(isPostQrData.count(1))
# 				isPostDataShiftB.append(isPostQrData)
# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_abhoypur_shiftB:
# 					if t == (22,00,00,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					elif patrol_abhoypur_shiftB.index(t) > patrol_abhoypur_shiftB.index((22,00,00,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)

# 				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftB.append(isPatrolQrData)

# 		if siteid == "Emami-Pacharia":
# 			if qrid[1] == 'QR':
# 				for t in post_shiftB_ap:
# 					if t == (23,30,00,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					elif post_shiftB_ap.index(t) > post_shiftB_ap.index((23,30,00,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)

# 				PostunAttendedShiftB.append(isPostQrData.count(0))
# 				PostAttendedShiftB.append(isPostQrData.count(1))
# 				isPostDataShiftB.append(isPostQrData)
			
# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_pacharia_shiftB:
# 					# if t == (21,00,00,00):
# 					if t == (23,00,2,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					# elif patrol_pacharia_shiftB.index(t) > patrol_pacharia_shiftB.index((21,00,00,00)):
# 					elif patrol_pacharia_shiftB.index(t) > patrol_pacharia_shiftB.index((23,00,2,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)


# 				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftB.append(isPatrolQrData)

# 		if siteid == "Emami-Amingaon":
# 			if qrid[1] == 'QR':
# 				for t in post_shiftB_ag:
# 					if t == (23,30,00,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					elif post_shiftB_ag.index(t) > post_shiftB_ag.index((23,30,00,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPostQrData.append(1)
# 						else:
# 							isPostQrData.append(0)

# 				PostunAttendedShiftB.append(isPostQrData.count(0))
# 				PostAttendedShiftB.append(isPostQrData.count(1))
# 				isPostDataShiftB.append(isPostQrData)

# 			if qrid[1] == 'QR PATROL':
# 				for t in patrol_amingaon_shiftB:
# 					if t == (23,00,1,00):
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					elif patrol_amingaon_shiftB.index(t) > patrol_amingaon_shiftB.index((23,00,1,00)):
# 						start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)
# 					else:
# 						start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
# 						end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
# 						start=tz.localize(start)
# 						end=tz.localize(end)
# 						data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
# 						if data:
# 							isPatrolQrData.append(1)
# 							patrolCount.append(len(data))
# 						else:
# 							isPatrolQrData.append(0)
# 							patrolCount.append(0)

# 				PatrolunAttendedShiftB.append(isPatrolQrData.count(0))
# 				PatrolAttendedShiftB.append(isPatrolQrData.count(1))
# 				isPatrolDataShiftB.append(isPatrolQrData)

# 	PostmaxAShiftA = max(PostAttendedShiftA)
# 	PostmaxUShiftA = max(PostunAttendedShiftA)
# 	PatrolmaxAShiftA = max(PatrolAttendedShiftA)
# 	PatrolmaxUShiftA = max(PatrolunAttendedShiftA)

# 	PostmaxAShiftB = max(PostAttendedShiftB)
# 	PostmaxUShiftB = max(PostunAttendedShiftB)
# 	PatrolmaxAShiftB = max(PatrolAttendedShiftB)
# 	PatrolmaxUShiftB = max(PatrolunAttendedShiftB)
# 	# print()
# 	# print("-------------------------------------------------")
# 	# print(place_list_siteid)
# 	# print("-----------------All Post---------------------")
# 	# print(allPost)
# 	# print("-----------------------All Patrol--------------")
# 	# print(allPatrol)
# 	maxLen = int(max(lenList))
# 	# BASE_DIR = Path(__file__).resolve().parent.parent
# 	module_dir = os.path.dirname(__file__)
# 	Postfile = ''
# 	Patrolfile = ''

# 	# filename = os.path.join(module_dir,"Post Monitoring Report Format_abhu.xlsx")


# 	if siteid == "Emami-Abhoypur":
# 		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Abhoypur.xlsx'))
# 		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Abhoypur.xlsx'))
# 	if siteid == "Emami-Amingaon":
# 		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Amingaon.xlsx'))
# 		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Amingaon.xlsx'))
# 	if siteid == "Emami-Pacharia":
# 		Postfile = (os.path.join(module_dir, 'Post Monitoring Report Format - Pacharia.xlsx'))
# 		Patrolfile = (os.path.join(module_dir, 'Patrol Monitoring Report Format - Pacharia.xlsx'))
	
# 	wb1Post = xl.load_workbook(Postfile)
# 	wb1Patrol = xl.load_workbook(Patrolfile)
	  
# 	# opening the destination excel file 
# 	filename1 =siteid+" Post Monitoring Report.xlsx"
# 	filename2 =siteid+" Patrol Monitoring Report.xlsx"
# 	wb2Post = xl.load_workbook(Postfile)
# 	wb2Patrol = xl.load_workbook(Patrolfile)

# 	wsPostDayShift = wb2Post["Day Shift"]
# 	wsPostNightShift = wb2Post["Night Shift"]
# 	wsPatrolDayShift = wb2Patrol["Day Shift"]
# 	wsPatrolNightShift = wb2Patrol["Night Shift"]

# 	wsPostDayShift['A4'] = siteid
# 	wsPostNightShift['A4'] = siteid
# 	wsPatrolDayShift['A4'] = siteid
# 	wsPatrolNightShift['A4'] = siteid
	
# 	if siteid == "Emami-Abhoypur":
# 		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
# 		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
# 		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
# 		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
# 	if siteid == "Emami-Amingaon":
# 		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
# 		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
# 		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
# 		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
# 	if siteid == "Emami-Pacharia":
# 		wsPostDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
# 		wsPostNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
# 		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
# 		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
	
# 	finalPost = []
# 	# if siteid == 'Emami-Abhoypur':
# 	# 	for i in allPost:
# 	# 		if i not in ('Factory Front Side', 'Back Side Utility Area'):
# 	# 			finalPost.append(i)
# 	# elif siteid == 'Emami-Amingaon':
# 	# 	for i in allPost:
# 	# 		if i not in ('Back Side DG Area','Loading Point'):
# 	# 			finalPost.append(i)
# 	# else:
# 	# 	for i in allPost:
# 	# 		finalPost.append(i)

# 	if siteid == 'Emami-Abhoypur':
# 		for i in AllPost:
# 			if i not in ('Factory Front Side [Area-Outside(Emami-Abhoypur)]', 'Back Side Utility Area [Area-Outside(Emami-Abhoypur)]'):
# 			# if i not in ('Factory Front Side', 'Back Side Utility Area'):
# 			# if 'Factory Front Side' or 'Back Side Utility Area' not in i:
# 				finalPost.append(i)
# 	elif siteid == 'Emami-Amingaon':
# 		for i in AllPost:
# 			# if i not in ('Back Side DG Area','Loading Point'):
# 			if i not in ('Back Side DG Area [Area-Outside(Emami-Amingaon)]','Loading Point [Area-Outside(Emami-Amingaon)]'):
# 			# if 'Back Side DG Area' or 'Loading Point' not in i:
# 				finalPost.append(i)
# 	else:
# 		for i in AllPost:
# 			finalPost.append(i)

# 	thin = Side (border_style="thin", color="000000")

# 	shiftAcolDay = 3
# 	shiftAiday = 0
	
# 	for post in finalPost:
# 		shiftArowDPost = 19
# 		for data in isPostDataShiftA[shiftAiday]:
# 			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
# 			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 			if data:
# 				wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = '64c620', fill_type="solid")
# 				wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			else:
# 				wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
# 				wsPostDayShift.cell(row = shiftArowDPost, column = shiftAcolDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			shiftArowDPost += 1
# 		shiftAiday += 1
# 		shiftAcolDay += 1


# 	shiftAcol = 3
# 	shiftAi = 0
# 	for post in allPost:
# 		shiftArowNPost = 19
# 		for data in isPostDataShiftB[shiftAi]:
# 			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
# 			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 			if data:
# 				wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
# 				wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			else:
# 				wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
# 				wsPostNightShift.cell(row = shiftArowNPost, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			shiftArowNPost += 1
# 		shiftAi += 1
# 		shiftAcol += 1

# 	shiftBcol = 3
# 	shiftBi = 0
# 	for post in allPatrol:
# 		shiftBrowDPatrol = 19
# 		shiftBroNPatrol = 19
# 		for data in isPatrolDataShiftA[shiftBi]:
# 			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
# 			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 			if data:
# 				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
# 				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			else:
# 				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
# 				wsPatrolDayShift.cell(row = shiftBrowDPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			shiftBrowDPatrol += 1
# 		for data in isPatrolDataShiftB[shiftBi]:
# 			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
# 			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 			if data:
# 				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
# 				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			else:
# 				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
# 				wsPatrolNightShift.cell(row = shiftBroNPatrol, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 			shiftBroNPatrol += 1

# 		shiftBi += 1
# 		shiftBcol += 1

# 	col = 3
# 	for post in range(0, len(allPost)):
# 		wsPostNightShift.cell(row = 18, column = col, value=AllPost[post])
# 		wsPostNightShift.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 		wsPostNightShift.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPostNightShift.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
# 		wsPostNightShift.column_dimensions[get_column_letter(col)].width = 25
# 		wsPostNightShift.row_dimensions[18].height = 70
# 		col += 1
	
# 	colDay = 3
# 	for post in range(0, len(finalPost)):
# 		wsPostDayShift.cell(row = 18, column = colDay, value=finalPost[post])
# 		wsPostDayShift.cell(row = 18, column = colDay).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 		wsPostDayShift.cell(row = 18, column = colDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPostDayShift.cell(row = 18, column = colDay).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

# 		wsPostDayShift.column_dimensions[get_column_letter(colDay)].width = 25
# 		wsPostDayShift.row_dimensions[18].height = 70
# 		colDay += 1
	
# 	Patrolcol = 3
# 	for post in range(0, len(allPatrol)):
# 		wsPatrolDayShift.cell(row = 18, column = Patrolcol, value=AllPatrol[post])
# 		wsPatrolNightShift.cell(row = 18, column = Patrolcol, value=AllPatrol[post])
		
# 		wsPatrolDayShift.cell(row = 18, column = Patrolcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 		wsPatrolNightShift.cell(row = 18, column = Patrolcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
# 		wsPatrolDayShift.cell(row = 18, column = Patrolcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPatrolNightShift.cell(row = 18, column = Patrolcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)

# 		wsPatrolDayShift.cell(row = 18, column = Patrolcol).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
# 		wsPatrolNightShift.cell(row = 18, column = Patrolcol).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

# 		wsPatrolDayShift.column_dimensions[get_column_letter(Patrolcol)].width = 25
# 		wsPatrolNightShift.column_dimensions[get_column_letter(Patrolcol)].width = 25
# 		wsPatrolDayShift.row_dimensions[18].height = 70
# 		wsPatrolNightShift.row_dimensions[18].height = 70
# 		Patrolcol += 1

# 	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
# 	wsPostDayShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = colDay-1)
# 	wsPostNightShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	
# 	wsPatrolDayShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = Patrolcol-1)
# 	wsPatrolNightShift.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = Patrolcol-1)
	
# 	# ws1['C17'] = 'Post Name'
# 	wsPostDayShift['C17'] = 'Post Name'
# 	wsPostNightShift['C17'] = 'Post Name'
# 	wsPatrolDayShift['C17'] = 'Post Name'
# 	wsPatrolNightShift['C17'] = 'Post Name'
# 	# ws1['C17'].font = Font(bold=True)
# 	wsPostDayShift['C17'].font = Font(bold=True)
# 	wsPostNightShift['C17'].font = Font(bold=True)
# 	wsPatrolDayShift['C17'].font = Font(bold=True)
# 	wsPatrolNightShift['C17'].font = Font(bold=True)

# 	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
# 	wsPostDayShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
# 	wsPostNightShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
# 	wsPatrolDayShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
# 	wsPatrolNightShift.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	
# 	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 	wsPostDayShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 	wsPostNightShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 	wsPatrolDayShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# 	wsPatrolNightShift.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 	bordercolPost = 3
# 	for post in range(0, len(allPost)):
# 		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPostNightShift.cell(row = 17, column = bordercolPost).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		bordercolPost += 1
	
# 	bordercolPostDay = 3
# 	for post in range(0, len(finalPost)):
# 		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPostDayShift.cell(row = 17, column = bordercolPostDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		bordercolPostDay += 1

# 	bordercolPatrol = 3
# 	for post in range(0, len(allPatrol)):
# 		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPatrolDayShift.cell(row = 17, column = bordercolPatrol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		wsPatrolNightShift.cell(row = 17, column = bordercolPatrol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
# 		bordercolPatrol += 1

# 	# ws1['E8'] = place_list_siteid[unAttendedShiftA.index(maxUShiftA)]
# 	# ws1['E9'] = place_list_siteid[AttendedShiftA.index(maxAShiftA)]
# 	# ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
# 	# ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
# 	# ws1.column_dimensions['F'].auto_size = True
# 	wsPostDayShift['E8'] = allPost[PostunAttendedShiftA.index(PostmaxUShiftA)]
# 	wsPostDayShift['E9'] = allPost[PostAttendedShiftA.index(PostmaxAShiftA)]
# 	wsPostDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPostDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPostNightShift['E8'] = allPost[PostunAttendedShiftB.index(PostmaxUShiftB)]
# 	wsPostNightShift['E9'] = allPost[PostAttendedShiftB.index(PostmaxAShiftB)]
# 	wsPostNightShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPostNightShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPatrolDayShift['E8'] = allPatrol[PatrolunAttendedShiftA.index(PatrolmaxUShiftA)]
# 	wsPatrolDayShift['E9'] = allPatrol[PatrolAttendedShiftA.index(PatrolmaxAShiftA)]
# 	wsPatrolDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPatrolDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPatrolNightShift['E8'] = allPatrol[PatrolunAttendedShiftB.index(PatrolmaxUShiftB)]
# 	wsPatrolNightShift['E9'] = allPatrol[PatrolAttendedShiftB.index(PatrolmaxAShiftB)]
# 	wsPatrolNightShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPatrolNightShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

# 	# ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
# 	# ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPostDayShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPostDayShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPostNightShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPostNightShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPatrolDayShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPatrolDayShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

# 	wsPatrolNightShift['A18'].alignment = Alignment(horizontal='center', vertical='center')
# 	wsPatrolNightShift['B18'].alignment = Alignment(horizontal='center', vertical='center')

# 	# saving the destination excel file
# 	wb2Post.save(str(filename1))
# 	wb2Patrol.save(str(filename2))

# 	sitenamemail=siteid

# 	try:
# 		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
# 		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

# 		# print(df)

# 		df_1=df.loc[df['siteid']==siteid]
# 		mail_list=df_1.iloc[:,1:].values.tolist()
# 		mail_list=[element for element in mail_list[0] if '@' in str(element)]
# 		send_to_list=mail_list
# 		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
# 		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
# 		subject=f"{siteid}- Daily Remote Monitoring Report"
# 		try:
# 			body=f"Dear Sir,\n\nPlease find attached the automated daily Post & Patrol Monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
# 		except:
# 			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
# 		# html_message = render_to_string('table/table.html')
# 		# text_content=strip_tags(html_message)
# 		# message = get_template('table/table_1.html').render()
# 		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
# 		# email.content_subtype ="html"# Main content is now text/html

# 	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
# 	# 	# email.attach_alternative(html_message,"text/html")

# 		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
# 		mail.attach_file(filename1)
# 		mail.attach_file(filename2)
# 		mail.send()
# 		# print("done!")
# 		# return HttpResponse("Hello World", content_type="text/plain")
# 	except:
# 		send_to_list=['arnab@achillesresolute.com']
# 		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
# 		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
# 		mail.send()

	

# 	return HttpResponse("Success")



	# if siteid=="Teloijan-Tea":
	# 	sitenamemail="Poddar House(Teloijan-Tea)"
	# else:
	# 	sitenamemail=siteid


	# try:
	# 	print("program is here")
	# 	mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
	# 	df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

	# 	# print(df)

	# 	df_1=df.loc[df['siteid']==siteid]
	# 	mail_list=df_1.iloc[:,1:].values.tolist()
	# 	mail_list=[element for element in mail_list[0] if '@' in str(element)]
	# 	send_to_list=mail_list
	# 	#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
	# 	# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
	# 	subject=f"{siteid}- Daily Post Monitoring Report"
	# 	try:
	# 		body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
	# 	except:
	# 		body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
	# 	# html_message = render_to_string('table/table.html')
	# 	# text_content=strip_tags(html_message)
	# 	# message = get_template('table/table_1.html').render()
	# 	# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.content_subtype ="html"# Main content is now text/html

	# # 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# # 	# email.attach_alternative(html_message,"text/html")

	# 	mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
	# 	mail.attach_file(filename1)
	# 	mail.send()
	# 	# print("done!")
	# 	# return HttpResponse("Hello World", content_type="text/plain")
	# except:
	# 	send_to_list=['arnab@achillesresolute.com']
	# 	body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
	# 	mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
	# 	mail.send()


	# return HttpResponse("Success")

def qrReportabg(request):
	siteid = request.GET.get('sitecode',"")

	if siteid=="Style Baazar":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]

		shiftB = [(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]

		shiftG=[(10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]


	elif siteid=="Bazaar Kolkata BZKT 01":
		shiftA = [(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00),(8,00,9,00)]

		shiftG=[(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]

	elif siteid=="Parklane Hotel Resorts RNHL01":
		shiftA = [(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00),(8,00,9,00)]

		shiftG=[(10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00),(22,00,23,00)]


	elif siteid=="Sonotel Hotel and Resort OZNE04":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		# shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]

		shiftG=[(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00),(5,00,6,00)]

		# shiftM=[(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]

	elif siteid=="Colliers International Property Services CLIS01":
		shiftA = [(1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00)]

		# shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftB = [(13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00)]

		shiftG=[(10,00,11,00), (11,00,12,00),(12,00,13,00),(13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]




	else:
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]

		shiftB = [(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]

		shiftG=[(10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]
		

	allPlace = []
	uuidListA = []
	uuidListB = []
	# uuidListC = []
	uuidListG = []
	place_list_siteidA = []
	place_list_siteidB = []
	# place_list_siteidC = []
	place_list_siteidG = []


	timeshiftA=[]
	timeshiftB=[]
	timeshiftG=[]

	for ele_a in shiftA:
		if ele_a[1]<10:
			b1="0"+str(ele_a[1])
		else:
			b1=str(ele_a[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a[3])
		else:
			d1=str(ele_a[3])
		 


		timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	for ele_b in shiftB:
		if ele_b[1]<10:
			b1="0"+str(ele_b[1])
		else:
			b1=str(ele_b[1])

		if ele_b[3]<10:
			d1="0"+str(ele_b[3])
		else:
			d1=str(ele_b[3])
		timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)

	for ele_g in shiftG:
		if ele_g[1]<10:
			b1="0"+str(ele_g[1])
		else:
			b1=str(ele_g[1])

		if ele_g[3]<10:
			d1="0"+str(ele_g[3])
		else:
			d1=str(ele_g[3])
		timeshiftG.append(str(ele_g[0])+":"+b1+"-"+str(ele_g[2])+":"+d1)


	# print(timeshiftB)




	if siteid=="Style Baazar":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="Bazaar Kolkata BZKT 01" or siteid=="Colliers International Property Services CLIS01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))
		# print(df_A)
		# print(df_B)
		# print(df_G)

	elif siteid=="Sonotel Hotel and Resort OZNE04":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AC")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AC")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="M",placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="Parklane Hotel Resorts RNHL01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AC")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AC")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))




	else:
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))


	if siteid=="Bazaar Kolkata BZKT 01":


		uuidListA=df_A["uuid"].tolist()
		df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
		place_list_siteidA1=df_A["fullname"].tolist()
		place_list_siteidA=[]
		for ele in place_list_siteidA1:
			ele1=ele[:len(ele)-(len(siteid)+3)]
			place_list_siteidA.append(ele1)





		uuidListB=df_B["uuid"].tolist()
		df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
		place_list_siteidB1=df_B["fullname"].tolist()

		place_list_siteidB=[]
		for ele in place_list_siteidB1:
			ele1=ele[:len(ele)-(len(siteid)+3)]
			place_list_siteidB.append(ele1)



		# uuidListC=df_C["uuid"].tolist()
		# df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
		# place_list_siteidC=df_C["fullname"].tolist()

		uuidListG=df_G["uuid"].tolist()
		df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
		place_list_siteidG1=df_G["fullname"].tolist()

		place_list_siteidG=[]
		for ele in place_list_siteidG1:
			ele1=ele[:len(ele)-(len(siteid)+3)]
			place_list_siteidG.append(ele1)

	else:
		uuidListA=df_A["uuid"].tolist()
		df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
		place_list_siteidA=df_A["fullname"].tolist()

		uuidListB=df_B["uuid"].tolist()
		df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
		place_list_siteidB=df_B["fullname"].tolist()

		# uuidListC=df_C["uuid"].tolist()
		# df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
		# place_list_siteidC=df_C["fullname"].tolist()

		uuidListG=df_G["uuid"].tolist()
		df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
		place_list_siteidG=df_G["fullname"].tolist()



	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)


	if siteid=="":
		isDataShiftG = []
		AttendedShiftG = []
		unAttendedShiftG = []

		for qrid in uuidListG:
			isQrData = []
			for t in shiftG:
				if t == (23,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftG.index(t) > shiftG.index((23,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftG.append(isQrData.count(0))
			AttendedShiftG.append(isQrData.count(1))
			isDataShiftG.append(isQrData)


	else:


		isDataShiftG = []
		AttendedShiftG = []
		unAttendedShiftG = []
		for qrid in uuidListG:
			isQrData = []
			for t in shiftG:
				
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftG.append(isQrData.count(0))
			AttendedShiftG.append(isQrData.count(1))
			isDataShiftG.append(isQrData)







	isDataShiftB = []
	AttendedShiftB = []
	unAttendedShiftB = []

	for qrid in uuidListB:
		isQrData = []
		for t in shiftB:
			if t == (23,00,00,00):
				# print("Inside If")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			elif shiftB.index(t) > shiftB.index((23,00,00,00)):
				# print("Inside Elif")
				# print(t)
				start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			else:
				# print("Else Block")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
		unAttendedShiftB.append(isQrData.count(0))
		AttendedShiftB.append(isQrData.count(1))
		isDataShiftB.append(isQrData)


	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	# maxAShiftC = max(AttendedShiftC)
	# maxUShiftC = max(unAttendedShiftC)
	maxAShiftG = max(AttendedShiftG)
	maxUShiftG = max(unAttendedShiftG)
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)


	if siteid!="Himadri Chemical Ltd.":
		filename=os.path.join(module_dir,"Patrol Monitoring Report Format Two Shift_G.xlsx")
	else:
		filename = os.path.join(module_dir,"Patrol Monitoring Report Format Two Shift_G.xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	if siteid=="Apollo Hospitals":
		filename1 =siteid+" Patrol Monitoring Report.xlsx"
		#filename1 =siteid+" Post Monitoring Report.xlsx"
	elif siteid=="Bazaar Kolkata BZKT 01":
		filename1 ="Metro Retails Pvt Ltd"+"_Post Monitoring Report.xlsx"

	else:
		filename1 =siteid+" Post Monitoring Report.xlsx"
		#filename1 =siteid+" Patrol Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Shift A"]
	ws2 = wb2["Shift B"]
	ws3=wb2["Shift G"]
	# ws3 = wb2["Shift C"]
	# ws4 = wb2["Shift G"]
	if siteid=="Bazaar Kolkata BZKT 01":
		ws1['A4'] = "Metro Retails Pvt Ltd"
		ws2['A4'] = "Metro Retails Pvt Ltd"
		ws3['A4'] = "Metro Retails Pvt Ltd"


	else:
		ws1['A4'] = siteid
		ws2['A4'] = siteid
		ws3['A4'] = siteid
	# ws4['A4'] = siteid
	if siteid=="Apollo Hospitals":
		ws1['A3']="Patrol Monitoring System"
		ws2['A3']="Patrol Monitoring System"

	if siteid=="Style Baazar":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"
	elif siteid=="Bazaar Kolkata BZKT 01":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	elif siteid=="Sonotel Hotel and Resort OZNE04":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift M :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"

	elif siteid=="Parklane Hotel Resorts RNHL01":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 09:00 Hours"
		ws3['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 11:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 23:00 Hours"

	elif siteid=="Colliers International Property Services CLIS01":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 01:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 13:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 13:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 01:00 Hours"
		ws3['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"





	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 10:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"


	rownum1=19
	thin = Side (border_style="thin", color="000000")

	for ele_ta in timeshiftA:
		ele_ta1=ele_ta.split("-")
		ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
		ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
		ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum1=rownum1+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1


	rownum2=19
	for ele_tb in timeshiftB:
		ele_tb1=ele_tb.split("-")
		ws2.cell(row=rownum2,column=1).value=ele_tb1[0]
		ws2.cell(row=rownum2,column=2).value=ele_tb1[1]
		ws2.cell(row=rownum2,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row=rownum2,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum2=rownum2+1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1


	rownum3=19
	for ele_tg in timeshiftG:
		ele_tb1=ele_tg.split("-")
		ws3.cell(row=rownum3,column=1).value=ele_tb1[0]
		ws3.cell(row=rownum3,column=2).value=ele_tb1[1]
		ws3.cell(row=rownum3,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row=rownum3,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum3=rownum3+1

	shiftCcol = 3
	shiftCi = 0
	for postC in place_list_siteidG:
		shiftCrow = 19
		for data in isDataShiftG[shiftCi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftCrow += 1
		shiftCi += 1
		shiftCcol += 1


	# shiftGcol = 3
	# shiftGi = 0
	# for postG in place_list_siteidG:
	# 	shiftGrow = 19
	# 	for data in isDataShiftG[shiftGi]:
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		shiftGrow += 1
	# 	shiftGi += 1
	# 	shiftGcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteidG[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws3['C17'] = 'Post Name'

	ws3['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	# col = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
	# 	ws4.cell(row = 18, column = col, value=place_list_siteidG[post])
		
	# 	# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 	ws4.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
	# 	# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

	# 	# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
	# 	ws4.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

	# 	# ws1.column_dimensions[get_column_letter(col)].width = 25
	# 	ws4.column_dimensions[get_column_letter(col)].width = 25
	# 	# ws1.row_dimensions[18].height = 70
	# 	ws4.row_dimensions[18].height = 70
	# 	col += 1

	# # ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	# ws4.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	# ws4['C17'] = 'Post Name'

	# ws4['C17'].font = Font(bold=True)

	# # ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws4.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# # ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# ws4.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	# bordercol = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	bordercol += 1



	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	ws3['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	# ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	# ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')

	wb2.save(str(filename1))

	if siteid=="Bazaar Kolkata BZKT 01":
		sitenamemail="Metro Retails Pvt Ltd"
	else:
		sitenamemail=siteid


	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{sitenamemail}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()


	return HttpResponse("Success")



def qrReportpostpatrol(request):
	siteid = request.GET.get('sitecode',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="Supreme Industries":

		shiftA = [(8,00,8,30), (8,30,9,00), (9,00,9,30), (9,30,10,00),(10,00,10,30), (10,30,11,00), (11,00,11,30), (11,30,12,00), (12,00,12,30),(12,30,13,00), (13,00,13,30), (13,30,14,00),(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
			(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00),(19,00,19,30), (19,30,20,00)]

		#shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]
		shiftB=[(20,00,20,30),(20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00)]

		shiftA1=[(8,00,10,00),(10,00,12,00),(12,00,14,00),(14,00,16,00),(16,00,18,00),(18,00,20,00)]



	else:
		shiftA = [(8,00,8,30), (8,30,9,00), (9,00,9,30), (9,30,10,00),(10,00,10,30), (10,30,11,00), (11,00,11,30), (11,30,12,00), (12,00,12,30),(12,30,13,00), (13,00,13,30), (13,30,14,00),(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
			(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00),(19,00,19,30), (19,30,20,00)]

		#shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]
		shiftB=[(20,00,20,30),(20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00)]

		shiftA1=[(8,00,10,00),(10,00,12,00),(12,00,14,00),(14,00,16,00),(16,00,18,00),(18,00,20,00)]




	allPlace = []
	uuidListA = []
	uuidListB = []

	uuidListA1=[]
	# uuidListC = []
	# uuidListG = []
	place_list_siteidA = []
	place_list_siteidB = []

	place_list_siteidA1=[]
	# place_list_siteidC = []
	# place_list_siteidG = []


	timeshiftA=[]
	timeshiftB=[]

	timeshiftA1=[]
	timeshiftB1=[]

	for ele_a in shiftA:
		if ele_a[1]<10:
			b1="0"+str(ele_a[1])
		else:
			b1=str(ele_a[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a[3])
		else:
			d1=str(ele_a[3])
		 


		timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	for ele_b in shiftB:
		if ele_b[1]<10:
			b1="0"+str(ele_b[1])
		else:
			b1=str(ele_b[1])

		if ele_b[3]<10:
			d1="0"+str(ele_b[3])
		else:
			d1=str(ele_b[3])
		timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)


	for ele_a1 in shiftA1:
		if ele_a1[1]<10:
			b1="0"+str(ele_a1[1])
		else:
			b1=str(ele_a1[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a1[3])
		else:
			d1=str(ele_a1[3])
		 


		timeshiftA1.append(str(ele_a1[0])+":"+b1+"-"+str(ele_a1[2])+":"+d1)


	if siteid=="Supreme Industries":

		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid,uuidtype="QR").values()))

		df_A1=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid,uuidtype="QR PATROL").values()))

		uuidListA=df_A["uuid"].tolist()
		df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
		place_list_siteidA=df_A["fullname"].tolist()

		uuidListB=df_B["uuid"].tolist()
		df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
		place_list_siteidB=df_B["fullname"].tolist()

		uuidListA1=df_A1["uuid"].tolist()
		df_A1["fullname"]=df_A1["name"]+"-"+df_A1["placeid_id"]
		place_list_siteidA1=df_A1["fullname"].tolist()



	else:
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_A1=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name__icontains=siteid,uuidtype="QR PATROL").values()))

		uuidListA=df_A["uuid"].tolist()
		df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
		place_list_siteidA=df_A["fullname"].tolist()

		uuidListB=df_B["uuid"].tolist()
		df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
		place_list_siteidB=df_B["fullname"].tolist()

		uuidListA1=df_A1["uuid"].tolist()
		df_A1["fullname"]=df_A1["name"]+"-"+df_A1["placeid_id"]
		place_list_siteidA1=df_A1["fullname"].tolist()


	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	# shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
	# 			(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
	# 			(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)



	#patrolling
	isDataShiftA1 = []
	AttendedShiftA1 = []
	unAttendedShiftA1 = []



	# tz=timezone.get_current_timezone()

	for qrid in uuidListA1:
		isQrData = []
		for t in shiftA1:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA1.append(isQrData.count(0))
		AttendedShiftA1.append(isQrData.count(1))
		isDataShiftA1.append(isQrData)





	if siteid=="Supreme Industries":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:

			isQrData = []
			for t in shiftB:
				if t == (23,30,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,30,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)

			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)

	else:
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:

			isQrData = []
			for t in shiftB:
				if t == (23,30,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,30,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)

			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)



	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	maxAShiftA1 = max(AttendedShiftA1)
	maxUShiftA1 = max(unAttendedShiftA1)
	# maxAShiftC = max(AttendedShiftC)
	# maxUShiftC = max(unAttendedShiftC)
	# maxAShiftG = max(AttendedShiftG)
	# maxUShiftG = max(unAttendedShiftG)
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)

	filename=os.path.join(module_dir,"both.xlsx")

	# if siteid!="Himadri Chemical Ltd.":
	# 	filename=os.path.join(module_dir,"Patrol Monitoring Report Format Two Shift.xlsx")
	# else:
	# 	filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime1.xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)

	filename1 =siteid+" Remote Monitoring Report.xlsx"
	  
	# opening the destination excel file 

	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Post Shift A"]
	ws2 = wb2["Post Shift B"]
	ws3 = wb2["Patrol Shift A"]
	# ws4 = wb2["Shift G"]
	ws1['A4'] = siteid
	ws2['A4'] = siteid
	ws3['A4'] = siteid
	# ws4['A4'] = siteid

	ws3['A3']="Patrol Monitoring System"
	# if siteid=="Apollo Hospitals":
	# 	ws1['A3']="Patrol Monitoring System"
	# 	ws2['A3']="Patrol Monitoring System"


	if siteid=="Supreme Industries":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"
	elif siteid=="The Methodist School":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
		ws3['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		ws3['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"



	rownum1=19
	thin = Side (border_style="thin", color="000000")

	for ele_ta in timeshiftA:
		ele_ta1=ele_ta.split("-")
		ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
		ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
		ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum1=rownum1+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1


	rownum2=19
	for ele_tb in timeshiftB:
		ele_tb1=ele_tb.split("-")
		ws2.cell(row=rownum2,column=1).value=ele_tb1[0]
		ws2.cell(row=rownum2,column=2).value=ele_tb1[1]
		ws2.cell(row=rownum2,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row=rownum2,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum2=rownum2+1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1


	rownum3=19


	for ele_ta2 in timeshiftA1:
		ele_ta1=ele_ta2.split("-")
		ws3.cell(row=rownum3,column=1).value=ele_ta1[0]
		ws3.cell(row=rownum3,column=2).value=ele_ta1[1]
		ws3.cell(row=rownum3,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row=rownum3,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum3=rownum3+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA1:
		shiftArow = 19
		for data in isDataShiftA1[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1

	# shiftCcol = 3
	# shiftCi = 0
	# for postC in place_list_siteidC:
	# 	shiftCrow = 19
	# 	for data in isDataShiftC[shiftCi]:
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		shiftCrow += 1
	# 	shiftCi += 1
	# 	shiftCcol += 1


	# shiftGcol = 3
	# shiftGi = 0
	# for postG in place_list_siteidG:
	# 	shiftGrow = 19
	# 	for data in isDataShiftG[shiftGi]:
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		shiftGrow += 1
	# 	shiftGi += 1
	# 	shiftGcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidA1)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteidA1[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws3['C17'] = 'Patrolling Post Name'

	ws3['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA1)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	# col = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
	# 	ws4.cell(row = 18, column = col, value=place_list_siteidG[post])
		
	# 	# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 	ws4.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
	# 	# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

	# 	# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
	# 	ws4.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

	# 	# ws1.column_dimensions[get_column_letter(col)].width = 25
	# 	ws4.column_dimensions[get_column_letter(col)].width = 25
	# 	# ws1.row_dimensions[18].height = 70
	# 	ws4.row_dimensions[18].height = 70
	# 	col += 1

	# # ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	# ws4.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	# ws4['C17'] = 'Post Name'

	# ws4['C17'].font = Font(bold=True)

	# # ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws4.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# # ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# ws4.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	# bordercol = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	bordercol += 1



	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteidA1[unAttendedShiftA1.index(maxUShiftA1)]
	ws3['E9'] = place_list_siteidA1[AttendedShiftA1.index(maxAShiftA1)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	# ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	# ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['B18'].alignment = Alignment(horizontal='center', vertical='center')


	wb2.save(filename1)

	sitenamemail=siteid


	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Remote Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily Remote monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()


	return HttpResponse("Success")


def qrReportabc(request):
	siteid = request.GET.get('sitecode',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="Park Hospital" or siteid=="Sterling Pharmacy" or siteid=="MANGALAM VIDYANIKETAN" or siteid=="ANUPAMA BIONATURALS" or siteid=="TK SENS":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

	elif siteid=="Kharagpur Metal KRRI 01" or siteid=="Kharagpur Metal KRRI 02" or siteid=="Torsa Machines" or siteid=="Aditya Birla AFRL03" or siteid=="Apeejay School APJY01" or siteid=="East India Pharmaceutical Works EIPH01": 
		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00),(21,00,22,00)]

		shiftC = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]

	elif siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT02" or siteid=="YORK PRINT YPRT04":

		shiftA = [(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00),(8,00,8,30),(8,30,9,00),(9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00),
			(11,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00)]

		shiftB=[(14,00,14,30),(14,30,15,00),(15,00,15,30),(15,30,16,00),(16,00,16,30),(16,30,17,00),(17,00,17,30),
			(17,30,18,00),(18,00,18,30),(18,30,19,00),(19,00,19,30),(19,30,20,00),(20,00,20,30),(20,30,21,00),(21,00,21,30),(21,30,22,00)]


		shiftC=[(22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00)]

	elif siteid=="AMRI Hospitals A01701":
		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]




	else:
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

	allPlace = []
	uuidListA = []
	uuidListB = []
	uuidListC = []
	# uuidListG = []
	# uuidListM = []
	place_list_siteidA = []
	place_list_siteidB = []
	place_list_siteidC = []
	# place_list_siteidG = []
	# place_list_siteidM = []


	timeshiftA=[]
	timeshiftB=[]
	timeshiftC=[]
	# timeshiftG=[]
	# timeshiftM=[]

	for ele_a in shiftA:
		if ele_a[1]<10:
			b1="0"+str(ele_a[1])
		else:
			b1=str(ele_a[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a[3])
		else:
			d1=str(ele_a[3])
		 


		timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	for ele_b in shiftB:
		if ele_b[1]<10:
			b1="0"+str(ele_b[1])
		else:
			b1=str(ele_b[1])

		if ele_b[3]<10:
			d1="0"+str(ele_b[3])
		else:
			d1=str(ele_b[3])
		timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)

	for ele_c in shiftC:
		if ele_c[1]<10:
			b1="0"+str(ele_c[1])
		else:
			b1=str(ele_c[1])

		if ele_c[3]<10:
			d1="0"+str(ele_c[3])
		else:
			d1=str(ele_c[3])
		timeshiftC.append(str(ele_c[0])+":"+b1+"-"+str(ele_c[2])+":"+d1)

	if siteid=="Park Hospital" or siteid=="Sterling Pharmacy" or siteid=="MANGALAM VIDYANIKETAN" or siteid=="ANUPAMA BIONATURALS" or siteid=="TK SENS":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="WILO MATHER AND PLATT PUMPS ACPL01" or siteid=="MACKERTICH CONSULTANCY SERVICES MKRH01" or siteid=="Uma Poly Solutions UPLY01" or siteid=="Apeejay School APJY01" or siteid=="Aditya Birla AFRL03":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name=siteid).values()))

	elif siteid=="East India Pharmaceutical Works EIPH01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name=siteid).values()))





	else:
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))



	uuidListA=df_A["uuid"].tolist()
	df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
	place_list_siteidA=df_A["fullname"].tolist()

	uuidListB=df_B["uuid"].tolist()
	df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
	place_list_siteidB=df_B["fullname"].tolist()

	uuidListC=df_C["uuid"].tolist()
	df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
	place_list_siteidC=df_C["fullname"].tolist()

	# uuidListG=df_G["uuid"].tolist()
	# df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
	# place_list_siteidG=df_G["fullname"].tolist()

	# uuidListM=df_M["uuid"].tolist()
	# df_M["fullname"]=df_M["name"]+"-"+df_M["placeid_id"]
	# place_list_siteidM=df_M["fullname"].tolist()


	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)

	isDataShiftB = []
	AttendedShiftB = []
	unAttendedShiftB = []
	for qrid in uuidListB:
		isQrData = []
		for t in shiftB:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftB.append(isQrData.count(0))
		AttendedShiftB.append(isQrData.count(1))
		isDataShiftB.append(isQrData)



	if siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT02" or siteid=="YORK PRINT YPRT04":
		isDataShiftC = []
		AttendedShiftC = []
		unAttendedShiftC = []
		for qrid in uuidListC:
			isQrData = []
			for t in shiftC:
				if t == (23,30,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftC.index(t) > shiftC.index((23,30,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftC.append(isQrData.count(0))
			AttendedShiftC.append(isQrData.count(1))
			isDataShiftC.append(isQrData)

	else:



		isDataShiftC = []
		AttendedShiftC = []
		unAttendedShiftC = []
		for qrid in uuidListC:
			isQrData = []
			for t in shiftC:
				if t == (23,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftC.index(t) > shiftC.index((23,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftC.append(isQrData.count(0))
			AttendedShiftC.append(isQrData.count(1))
			isDataShiftC.append(isQrData)

	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	maxAShiftC = max(AttendedShiftC)
	maxUShiftC = max(unAttendedShiftC)
	# maxAShiftG = max(AttendedShiftG)
	# maxUShiftG = max(unAttendedShiftG)
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)


	filename = os.path.join(module_dir,"abc.xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	if siteid!="Apollo Hospitals":
		filename1 =siteid+" Post Monitoring Report.xlsx"
	else:
		filename1 =siteid+" Patrol Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Shift A"]
	ws2 = wb2["Shift B"]
	ws3=wb2["Shift C"]
	# ws3 = wb2["Shift C"]
	# ws4 = wb2["Shift G"]
	ws1['A4'] = siteid
	ws2['A4'] = siteid
	ws3['A4'] = siteid
	# ws4['A4'] = siteid
	if siteid=="Park Hospital" or siteid=="Sterling Pharmacy" or siteid=="MANGALAM VIDYANIKETAN" or siteid=="ANUPAMA BIONATURALS" or siteid=="TK SENS" or siteid=="WILO MATHER AND PLATT PUMPS ACPL01":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
	elif siteid=="Kharagpur Metal KRRI 01" or siteid=="Kharagpur Metal KRRI 02" or siteid=="Torsa Machines" or siteid=="YORK PRINT YPRT02" or siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04" or siteid=="Aditya Birla AFRL03" or siteid=="Apeejay School APJY01": 
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"



	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"


	rownum1=19
	thin = Side (border_style="thin", color="000000")

	for ele_ta in timeshiftA:
		ele_ta1=ele_ta.split("-")
		ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
		ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
		ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum1=rownum1+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1


	rownum2=19
	for ele_tb in timeshiftB:
		ele_tb1=ele_tb.split("-")
		ws2.cell(row=rownum2,column=1).value=ele_tb1[0]
		ws2.cell(row=rownum2,column=2).value=ele_tb1[1]
		ws2.cell(row=rownum2,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row=rownum2,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum2=rownum2+1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1


	rownum3=19
	for ele_tc in timeshiftC:
		ele_tb1=ele_tc.split("-")
		ws3.cell(row=rownum3,column=1).value=ele_tb1[0]
		ws3.cell(row=rownum3,column=2).value=ele_tb1[1]
		ws3.cell(row=rownum3,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row=rownum3,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum3=rownum3+1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidC:
		shiftBrow = 19
		for data in isDataShiftC[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteidC[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws3['C17'] = 'Post Name'

	ws3['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
	ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')




	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')


	# saving the destination excel file
	wb2.save(str(filename1))

	if siteid=="AMRI Hospitals A01701":
		sitenamemail="AMRI Hospitals Dakuria"
	else:

		sitenamemail=siteid

	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Remote Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily Remote monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	return HttpResponse("Success")







def qrReportabcgm(request):

	siteid = request.GET.get('sitecode',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="Meridian Medical Hospital":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

		shiftG=[(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00)]

		shiftM=[(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]





	elif siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

		shiftG=[(8,00,10,00),(10,00,12,00), (12,00,14,00), (14,00,16,00),(16,00,18,00), (18,00,20,00)]

		shiftM=[(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]

	else:
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]

		shiftC = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

		shiftG=[(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00)]

		shiftM=[(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]



	allPlace = []
	uuidListA = []
	uuidListB = []
	uuidListC = []
	uuidListG = []
	uuidListM = []
	place_list_siteidA = []
	place_list_siteidB = []
	place_list_siteidC = []
	place_list_siteidG = []
	place_list_siteidM = []


	timeshiftA=[]
	timeshiftB=[]
	timeshiftC=[]
	timeshiftG=[]
	timeshiftM=[]

	for ele_a in shiftA:
		if ele_a[1]<10:
			b1="0"+str(ele_a[1])
		else:
			b1=str(ele_a[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a[3])
		else:
			d1=str(ele_a[3])
		 


		timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	for ele_b in shiftB:
		if ele_b[1]<10:
			b1="0"+str(ele_b[1])
		else:
			b1=str(ele_b[1])

		if ele_b[3]<10:
			d1="0"+str(ele_b[3])
		else:
			d1=str(ele_b[3])
		timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)

	for ele_c in shiftC:
		if ele_c[1]<10:
			b1="0"+str(ele_c[1])
		else:
			b1=str(ele_c[1])

		if ele_c[3]<10:
			d1="0"+str(ele_c[3])
		else:
			d1=str(ele_c[3])
		timeshiftC.append(str(ele_c[0])+":"+b1+"-"+str(ele_c[2])+":"+d1)


	for ele_g in shiftG:
		if ele_g[1]<10:
			b1="0"+str(ele_g[1])
		else:
			b1=str(ele_g[1])

		if ele_g[3]<10:
			d1="0"+str(ele_g[3])
		else:
			d1=str(ele_g[3])
		timeshiftG.append(str(ele_g[0])+":"+b1+"-"+str(ele_g[2])+":"+d1)

	for ele_m in shiftM:
		if ele_m[1]<10:
			b1="0"+str(ele_m[1])
		else:
			b1=str(ele_m[1])

		if ele_b[3]<10:
			d1="0"+str(ele_m[3])
		else:
			d1=str(ele_m[3])
		timeshiftM.append(str(ele_m[0])+":"+b1+"-"+str(ele_m[2])+":"+d1)

	if siteid=="Meridian Medical Hospital":

		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))
		df_M=pd.DataFrame(list(uuidTbl.objects.filter(other="M",placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name=siteid,uuidtype="QR").values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name=siteid,uuidtype="QR").values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name=siteid,uuidtype="QR PATROL").values()))
		df_M=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name=siteid,uuidtype="QR PATROL").values()))


	else:
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))
		df_C=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))
		df_G=pd.DataFrame(list(uuidTbl.objects.filter(other="G",placeid__siteid__name__icontains=siteid).values()))
		df_M=pd.DataFrame(list(uuidTbl.objects.filter(other="M",placeid__siteid__name__icontains=siteid).values()))


	uuidListA=df_A["uuid"].tolist()
	df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
	place_list_siteidA=df_A["fullname"].tolist()

	uuidListB=df_B["uuid"].tolist()
	df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
	place_list_siteidB=df_B["fullname"].tolist()

	uuidListC=df_C["uuid"].tolist()
	df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
	place_list_siteidC=df_C["fullname"].tolist()

	uuidListG=df_G["uuid"].tolist()
	df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
	place_list_siteidG=df_G["fullname"].tolist()

	uuidListM=df_M["uuid"].tolist()
	df_M["fullname"]=df_M["name"]+"-"+df_M["placeid_id"]
	place_list_siteidM=df_M["fullname"].tolist()


	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)

	isDataShiftB = []
	AttendedShiftB = []
	unAttendedShiftB = []
	for qrid in uuidListB:
		isQrData = []
		for t in shiftB:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftB.append(isQrData.count(0))
		AttendedShiftB.append(isQrData.count(1))
		isDataShiftB.append(isQrData)


	isDataShiftC = []
	AttendedShiftC = []
	unAttendedShiftC = []
	for qrid in uuidListC:
		isQrData = []
		for t in shiftC:
			if t == (23,00,00,00):
				# print("Inside If")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			elif shiftC.index(t) > shiftC.index((23,00,00,00)):
				# print("Inside Elif")
				# print(t)
				start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			else:
				# print("Else Block")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
		unAttendedShiftC.append(isQrData.count(0))
		AttendedShiftC.append(isQrData.count(1))
		isDataShiftC.append(isQrData)


	isDataShiftG = []
	AttendedShiftG = []
	unAttendedShiftG = []
	for qrid in uuidListG:
		isQrData = []
		for t in shiftG:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftG.append(isQrData.count(0))
		AttendedShiftG.append(isQrData.count(1))
		isDataShiftG.append(isQrData)


	isDataShiftM = []
	AttendedShiftM = []
	unAttendedShiftM = []
	for qrid in uuidListM:
		isQrData = []
		for t in shiftM:
			
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftM.append(isQrData.count(0))
		AttendedShiftM.append(isQrData.count(1))
		isDataShiftM.append(isQrData)





	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	maxAShiftC = max(AttendedShiftC)
	maxUShiftC = max(unAttendedShiftC)
	maxAShiftG = max(AttendedShiftG)
	maxUShiftG = max(unAttendedShiftG)
	maxAShiftM = max(AttendedShiftM)
	maxUShiftM = max(unAttendedShiftM)


	module_dir = os.path.dirname(__file__)


	if siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		filename = os.path.join(module_dir,"abcp.xlsx")
	else:

		filename = os.path.join(module_dir,"abcgm.xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	
	if siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		filename1 =siteid+" Remote Monitoring Report.xlsx"
	elif siteid!="Apollo Hospitals":
		filename1 =siteid+" Post Monitoring Report.xlsx"

	else:
		filename1 =siteid+" Patrol Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	if siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		ws1 = wb2["Shift A"]
		ws2 = wb2["Shift B"]
		ws3 = wb2["Shift C"]
		ws4 = wb2["Patrol"]
		ws5 = wb2["Shift M"]
	else:
		ws1 = wb2["Shift A"]
		ws2 = wb2["Shift B"]
		ws3 = wb2["Shift C"]
		ws4 = wb2["Shift G"]
		ws5 = wb2["Shift M"]


	ws1['A4'] = siteid
	ws2['A4'] = siteid
	ws3['A4'] = siteid
	ws4['A4'] = siteid
	ws5['A4'] = siteid


	if siteid=="Meridian Medical Hospital":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
		ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		ws5['A5'] = "Shift M :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	elif siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
		ws4['A5'] = "Patrol Monitoring :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws5['A5'] = "Shift M :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"

	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
		ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"
		ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		ws5['A5'] = "Shift M :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"


	thin = Side (border_style="thin", color="000000")

	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	shiftCcol = 3
	shiftCi = 0
	for postC in place_list_siteidC:
		shiftCrow = 19
		for data in isDataShiftC[shiftCi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftCrow += 1
		shiftCi += 1
		shiftCcol += 1


	shiftGcol = 3
	shiftGi = 0
	for postG in place_list_siteidG:
		shiftGrow = 19
		for data in isDataShiftG[shiftGi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftGrow += 1
		shiftGi += 1
		shiftGcol += 1

	shiftGcol = 3
	shiftGi = 0
	for postM in place_list_siteidM:
		shiftGrow = 19
		for data in isDataShiftM[shiftGi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws5.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws5.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws5.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws5.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftGrow += 1
		shiftGi += 1
		shiftGcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws3.cell(row = 18, column = col, value=place_list_siteidC[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws3.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws3.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws3['C17'] = 'Post Name'

	ws3['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidC)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws4.cell(row = 18, column = col, value=place_list_siteidG[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws4.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws4.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws4.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws4.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	if siteid=="AMRI-Bhubaneswar":
		ws4.merge_cells(start_row = 16, end_row=16, start_column=3, end_column = col-1)

		ws4['C16'] = 'Post Availability Timing & Post Name'

		ws4['C16'].font = Font(bold=True)

		# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		ws4.cell(row = 16, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 16, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

		timing_list=["08:00-21:00 hrs","09:00-21:00 hrs","10:00-19:00 hrs","09:00-21:00 hrs","09:00-18:00 hrs","08:00-20:00 hrs","06:00-18:00 hrs","09:00-21:00 hrs",
		"09:00-18:00 hrs","09:00-21:00 hrs","08:00-20:00 hrs","08:00-16:00 hrs","09:00-21:00 hrs","10:00-18:00 hrs","10:00-18:00 hrs"]

		rn=17
		cn=3
		k=0

		while k<len(timing_list):
			ws4.cell(row =rn, column = cn).value = timing_list[k]
			k=k+1
			cn+=1



	else:
		ws4.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

		ws4['C17'] = 'Post Name'

		ws4['C17'].font = Font(bold=True)

		# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		ws4.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
		# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws4.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidG)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws4.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1

	col = 3
	for post in range(0, len(place_list_siteidM)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws5.cell(row = 18, column = col, value=place_list_siteidM[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws5.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws5.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws5.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws5.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws5.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws5.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws5['C17'] = 'Post Name'

	ws5['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws5.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws5.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidM)):
		# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws5.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1



	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
	ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
	ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws5['E8'] = place_list_siteidM[unAttendedShiftM.index(maxUShiftM)]
	ws5['E9'] = place_list_siteidM[AttendedShiftM.index(maxAShiftM)]
	ws5['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws5['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws4['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws5['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws5['B18'].alignment = Alignment(horizontal='center', vertical='center')

	if siteid=="YORK PRINT YPRT03" or siteid=="YORK PRINT YPRT04":
		wb2.remove(wb2['Shift M'])
		wb2.save(str(filename1))
	else:
		wb2.save(str(filename1))


	# saving the destination excel file
	wb2.save(str(filename1))

	sitenamemail=siteid

	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Remote Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily Remote monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	return HttpResponse("Success")




def qrReporthouse(request):
	siteid = request.GET.get('sitecode',"")

	if siteid=="AMRI Hospital BBR":
		shiftA = [(6,00,6,30), (6,30,7,00),(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00), (9,00,9,30), (9,30,10,00),
			(10,00,10,30), (10,30,11,00), (11,00,11,30), (11,30,12,00), (12,00,12,30),
			(12,30,13,00), (13,00,13,30), (13,30,14,00),(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
			(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00),
			(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00),(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
			(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), (2,00,2,30),
			(2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30),
			(5,30,6,00)]

		uuidListA = []
		# uuidListB = []
		# uuidListC = []
		# uuidListG = []
		place_list_siteidA = []
		# place_list_siteidB = []
		# place_list_siteidC = []
		# place_list_siteidG = []


		timeshiftA=[]
		# timeshiftB=[]

		for ele_a in shiftA:
			if ele_a[1]<10:
				b1="0"+str(ele_a[1])
			else:
				b1=str(ele_a[1])

			if ele_a[3]<10:
				d1="0"+str(ele_a[3])
			else:
				d1=str(ele_a[3])
			 


			timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)


		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="G")|Q(other="ALL"),placeid__siteid__name=siteid,uuidtype="HOUSEKEEPING").values()))

		uuidListA=df_A["uuid"].tolist()
		df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
		place_list_siteidA=df_A["fullname"].tolist()


	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	# shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
	# 			(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
	# 			(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()



	for qrid in uuidListA:

		isQrData = []
		for t in shiftA:
			if t == (23,30,00,00):
				# print("Inside If")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			elif shiftA.index(t) > shiftA.index((23,30,00,00)):
				# print("Inside Elif")
				# print(t)
				start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
				# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			else:
				# print("Else Block")
				# print(t)
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)

		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)



	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	# maxAShiftB = max(AttendedShiftB)
	# maxUShiftB = max(unAttendedShiftB)
	# maxAShiftC = max(AttendedShiftC)
	# maxUShiftC = max(unAttendedShiftC)
	# maxAShiftG = max(AttendedShiftG)
	# maxUShiftG = max(unAttendedShiftG)
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)

	filename = os.path.join(module_dir,"housekeeping.xlsx")




	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	if siteid!="Apollo Hospitals":
		filename1 =siteid+" Monitoring Report.xlsx"
	else:
		filename1 =siteid+" Housekeeping Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["dayilyreport"]
	ws2 = wb2["TimeStamp"]
	# ws3 = wb2["Shift C"]
	# ws4 = wb2["Shift G"]
	ws1['A4'] = siteid
	# ws2['A4'] = siteid
	# ws3['A4'] = siteid
	# ws4['A4'] = siteid



	ws1['A5'] = "Day Report :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"

	
		
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"


	rownum1=19
	thin = Side (border_style="thin", color="000000")

	for ele_ta in timeshiftA:
		ele_ta1=ele_ta.split("-")
		ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
		ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
		ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum1=rownum1+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1



	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	# ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	# ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	# ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
	# ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
	# ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	# ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	# ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')

	time1=datetime.now()
	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time2_1+timedelta(hours=6)
	time2_11=time2_1-timedelta(days=1)

	start_time1=tz.localize(time2_11)
	end_time1=tz.localize(time2_1)

	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name=siteid,time__range=(start_time1,end_time1)).values()))

	if len(all_path)!=0:

		all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
		all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')


		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		# df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all=df_all.sort_values("time",ascending=True)
		df_all=df_all.reset_index()

		# print(df_all.columns)

		rownum=2

		for idx,rows in df_all.iterrows():
			ws2.cell(row = rownum, column = 1, value=rows['name'])
			ws2.cell(row = rownum, column = 2, value=rows['placeid_id'])
			ws2.cell(row = rownum, column = 3, value=rows['datetime1'])
			ws2.cell(row = rownum, column = 4, value=rows['empname'])

			ws2.cell(row = rownum, column = 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row = rownum, column = 2).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row = rownum, column = 3).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			ws2.cell(row = rownum, column = 4).border = Border(top=thin, left=thin, right=thin, bottom=thin)

			rownum=rownum+1



	wb2.save(str(filename1))


	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Housekeeping Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily housekeeping monitoring report pertaining to {siteid} for the date of {dt_1}.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily housekeeping monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()

	return HttpResponse("Success")
















		# for ele_b in shiftB:
		# 	if ele_b[1]<10:
		# 		b1="0"+str(ele_b[1])
		# 	else:
		# 		b1=str(ele_b[1])

		# 	if ele_b[3]<10:
		# 		d1="0"+str(ele_b[3])
		# 	else:
		# 		d1=str(ele_b[3])
		# 	timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)










def qrReportbhu2s(request):
	siteid = request.GET.get('sitecode',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="E-CITY Mall":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]

		shiftB = [(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]


	elif siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]
		shiftB=[(20,00,20,30),(20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00)]

	elif siteid=="Aum Packaging APKG01":
		shiftA = [(7,00,7,30),(7,30,8,00),(8,00,8,30),(8,30,9,00), (9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00),(11,00,11,30), (11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30),
				(13,30,14,00), (14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00),(16,00,16,30), (16,30,17,00), (17,00,17,30), (17,30,18,00),(18,00,18,30),(18,30,19,00),]

		shiftB=[(19,00,19,30),(19,30,20,00),(20,00,20,30),(20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00)]



		# print(len(shiftB))

	elif siteid=="Piramal Mahalakshmi":

		# shiftA = [(8,00,8,30),(8,30,9,00), (9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00),(11,00,11,30), (11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30),
		# 		(13,30,14,00), (14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00),(16,00,16,30), (16,30,17,00), (17,00,17,30), (17,30,18,00),(18,00,18,30),(18,30,19,00),
		# 		(19,00,19,30),(19,30,20,00)]

		shiftA=[(8,00,10,00), (10,00,12,00), (12,00,14,00), (14,00,16,00),(16,00,18,00), (18,00,20,00)]

		shiftB=[(20,00,22,00),(22,00,00,00), (00,00,2,00), (2,00,4,00), (4,00,6,00),(6,00,8,00)]


	elif siteid=="Apollo Hospitals":
		shiftA = [(8,00,11,00),(11,00,14,00),(14,00,17,00),(17,00,20,00)]

		shiftB = [(20,00,23,00),(23,00,2,00),(2,00,5,00),(5,00,8,00)]



	elif siteid=="Himadri Chemical Ltd.":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]
		shiftB=[(23,00,23,29), (23,30,00,00), (00,00,00,29), (00,30,00,59),(1,00,1,29), (1,30,1,59), (2,00,2,29), (2,30,2,59), (3,00,3,29), 
				(3,30,3,59), (4,00,4,29), (4,30,5,00)]

	elif siteid=="The Methodist School" or siteid=="Ahura Mazda" or siteid=="SRG Exim SEIM00001" or siteid=="Hathway Digital HWCC01" or siteid=="ACC Limited SIL ACCL01" or siteid=="Aum Packaging APKG01":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00)]

		shiftB = [(19,00,20,00),(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00)]

	elif siteid=="Manaksia Limited MNKS01" or siteid=="Spencer Retail SPRL10":
		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00)]

		shiftB = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),]


	elif siteid=="KONE ELEVATOR INDIA":
		shiftA = [(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]


	elif siteid=="Super Forging and Steel SSTS01" or siteid=="Mega Mall APAI01" or siteid=="Eco Jute EJUE02":
		shiftA = [ (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(20,00,21,00)]

		shiftB = [(21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00),(8,00,9,00)]




	else:

		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]

		shiftB = [(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]

	# print("Place List--->>>", place_list)
	# place_list_siteid = [i.name+' ('+siteid+')' for i in activeQR]
	
	
	allPlace = []
	uuidListA = []
	uuidListB = []
	# uuidListC = []
	# uuidListG = []
	place_list_siteidA = []
	place_list_siteidB = []
	# place_list_siteidC = []
	# place_list_siteidG = []


	timeshiftA=[]
	timeshiftB=[]

	for ele_a in shiftA:
		if ele_a[1]<10:
			b1="0"+str(ele_a[1])
		else:
			b1=str(ele_a[1])

		if ele_a[3]<10:
			d1="0"+str(ele_a[3])
		else:
			d1=str(ele_a[3])
		 


		timeshiftA.append(str(ele_a[0])+":"+b1+"-"+str(ele_a[2])+":"+d1)

	for ele_b in shiftB:
		if ele_b[1]<10:
			b1="0"+str(ele_b[1])
		else:
			b1=str(ele_b[1])

		if ele_b[3]<10:
			d1="0"+str(ele_b[3])
		else:
			d1=str(ele_b[3])
		timeshiftB.append(str(ele_b[0])+":"+b1+"-"+str(ele_b[2])+":"+d1)


	# print(timeshiftB)




	if siteid=="E-CITY Mall":

		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="Piramal Mahalakshmi" or siteid=="KISHORILAL SHYAMSUNDER" or siteid=="Vishakha Polyfab" or siteid=="Shri Sirisha Auto Pvt Ltd" or siteid=="TARSON" or siteid=="Ms S N ENTERPRISE" or siteid=="Artistery House" or siteid=="Maa Laxmi RELS 01" or siteid=="Maa Laxmi LINS 02" or siteid=="Karuna Management Services" or siteid=="Apeejay House" or siteid=="Small Business Fin Credit India" or siteid=="Small Business Fin Credit SBFC 02" or siteid=="Ms SP International" or siteid=="PARAS SURTI PRODUCTS" or siteid=="SNE BUSINESS INDIA" or siteid=="Suncity Residents" or siteid=="ALIPUR TERRACE":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))

	elif siteid=="Oriental Cusien" or siteid=="Global Motocorp LLP" or siteid=="WOW MOMO KASBA 01" or siteid=="WOW MOMO KASBA 02" or siteid=="Anamika Kala Sangam Trust" or siteid=="Mr Abhishek Sao" or siteid=="Tatwa Technologies TTWA01" or siteid=="ACC Limited ACCL02":
		# print("program is here")
		# print(siteid)
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))

		# print(df_A,df_B)

	elif siteid=="Thakur Prasad Sao" or siteid=="Todi House" or siteid=="Moore Height" or siteid=="Ahura Mazda" or siteid=="Kamal Mukherjee" or siteid=="Pioneer Packaging Industries" or siteid=="Calcutta Television Network CLTN00001" or siteid=="Calcutta Television Network CLTN00002" or siteid=="Sundarban Livestock SCEM01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))
	elif siteid=="J G Chemicals J003 02" or siteid=="Dharampal Sottopal" or siteid=="Ganesh Grains GHGS 00003" or siteid=="Tarsons Products TRSS01" or siteid=="ACC Ltd KOL ACCL01" or siteid=="SRG Exim SEIM00001" or siteid=="Maa Laxmi LINS 01" or siteid=="Vishva Electrotech Ltd VELI02" or siteid=="PAR Enterprises PRES 0001":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="B"),placeid__siteid__name=siteid).values()))


	elif siteid=="South Wind Project" or siteid=="Altamira Project" or siteid=="Associated Minerals" or siteid=="Aspira Loharuka Developers BGBN08" or siteid=="Masterly Kolkata Facility Maintenance CLIS02" or siteid=="Ganesh Grains GHGS 00002" or siteid=="Twenty One Century TOCP01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="C"),placeid__siteid__name=siteid).values()))

	elif siteid=="Great Eastern Stores APHL01" or siteid=="Kharagpur Metal Reforming Industries KRRI02" or siteid=="Kharagpur Metal Reforming Industries KRRI03" or siteid=="Plumbum Chemicals and Pigments PLCA01" or siteid=="PS PRIMARC PROJECTS LLP PRMP03" or siteid=="Chandrani Enterprises CNDR02":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="C"),placeid__siteid__name=siteid).values()))

	elif siteid=="ACC Limited ACCL01" or siteid=="Hathway Digital HWCC01" or siteid=="ACC Limited SIL ACCL01" or siteid=="Assam Hospital AMHL03" or siteid=="Aum Packaging APKG01" or siteid=="Aditya Birla AFRL02":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="ALL")|Q(other="AC")|Q(other="B")|Q(other="C"),placeid__siteid__name=siteid).values()))





	elif siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd." or siteid=="The Methodist School" or siteid=="KONE ELEVATOR INDIA" or siteid=="South Asian Red FM SHAL01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
	elif siteid=="Apollo Hospitals":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid,uuidtype="QR PATROL").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid,uuidtype="QR PATROL").values()))

	elif siteid=="Bengal GHG Nirman BDVL 03" or siteid=="Super Forging and Steel SSTS01" or siteid=="Wellside Developers WLLD01" or siteid=="Maxcab Wire and Cables MWSL01":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))

	elif siteid=="UDAYAN ENTREPRENEURS WELFARE UNWA01" or siteid=="Coastal Roadways CLLD01" or siteid=="Mega Mall APAI01" or siteid=="Eco Jute EJUE02" or siteid=="Manaksia Limited MNKS01" or siteid=="Spencer Retail SPRL10":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A")|Q(other="ALL"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B")|Q(other="ALL"),placeid__siteid__name__icontains=siteid,uuidtype="QR").values()))

	elif siteid=="PS PACE OWNERS ASSOCIATIONS":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A")|Q(other="ALL"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B")|Q(other="ALL")|Q(other="C"),placeid__siteid__name__icontains=siteid).values()))




	else:
		print("program is here")
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name=siteid).values()))


		



	uuidListA=df_A["uuid"].tolist()
	df_A["fullname"]=df_A["name"]+"-"+df_A["placeid_id"]
	place_list_siteidA=df_A["fullname"].tolist()

	uuidListB=df_B["uuid"].tolist()
	df_B["fullname"]=df_B["name"]+"-"+df_B["placeid_id"]
	place_list_siteidB=df_B["fullname"].tolist()

	# uuidListC=df_C["uuid"].tolist()
	# df_C["fullname"]=df_C["name"]+"-"+df_C["placeid_id"]
	# place_list_siteidC=df_C["fullname"].tolist()

	# uuidListG=df_G["uuid"].tolist()
	# df_G["fullname"]=df_G["name"]+"-"+df_G["placeid_id"]
	# place_list_siteidG=df_G["fullname"].tolist()





	# print(uuidlistA,place_list_siteidA)
	# print(df_A["name"])





	



	# print("allPlace--->>>", allPlace , len(allPlace))
	# postNo = []
	# for i in range(1, len(place_list_siteid)+1):
	# 	postNo.append('Post No.'+str(i))
	# lenList = [len(i) for i in place_list_siteid]
	# One Day Post Monitoring Report 10 to 10
	todayStr = date.today().strftime('%Y %d %m').split()
	today = date.today()
	yest = today - timedelta(days = 1)
	yestStr = yest.strftime('%Y %d %m').split()
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	# shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
	# 			(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
	# 			(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	isDataShiftA = []
	AttendedShiftA = []
	unAttendedShiftA = []

	tz=timezone.get_current_timezone()

	for qrid in uuidListA:
		isQrData = []
		for t in shiftA:
			# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
			start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
			end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
			
			start=tz.localize(start)
			end=tz.localize(end)
			data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
			if data:
				isQrData.append(1)
			else:
				isQrData.append(0)
		unAttendedShiftA.append(isQrData.count(0))
		AttendedShiftA.append(isQrData.count(1))
		isDataShiftA.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

	# isDataShiftB = []
	# AttendedShiftB = []
	# unAttendedShiftB = []
	# for qrid in uuidListB:
	# 	isQrData = []
	# 	for t in shiftB:
			
	# 		start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
	# 		end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
	# 		start=tz.localize(start)
	# 		end=tz.localize(end)
	# 		data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
	# 		if data:
	# 			isQrData.append(1)
	# 		else:
	# 			isQrData.append(0)
	# 	unAttendedShiftB.append(isQrData.count(0))
	# 	AttendedShiftB.append(isQrData.count(1))
	# 	isDataShiftB.append(isQrData)


	if siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd." or siteid=="Aum Packaging APKG01":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:

			isQrData = []
			for t in shiftB:
				if t == (23,30,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,30,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)

			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)


	elif siteid=="Piramal Mahalakshmi":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:

			isQrData = []
			for t in shiftB:
				if t == (22,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((22,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)

			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)



	elif siteid=="Apollo Hospitals":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:
			isQrData = []
			for t in shiftB:
				if t == (23,00,2,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,00,2,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)



	elif siteid=="KONE ELEVATOR INDIA":
		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		# tz=timezone.get_current_timezone()

		for qrid in uuidListB:
			isQrData = []
			for t in shiftB:
				# timeshiftA.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
				start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
				end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
				
				start=tz.localize(start)
				end=tz.localize(end)
				data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
				if data:
					isQrData.append(1)
				else:
					isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)




	else:


		isDataShiftB = []
		AttendedShiftB = []
		unAttendedShiftB = []

		for qrid in uuidListB:
			isQrData = []
			for t in shiftB:
				if t == (23,00,00,00):
					# print("Inside If")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				elif shiftB.index(t) > shiftB.index((23,00,00,00)):
					# print("Inside Elif")
					# print(t)
					start = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(todayStr[0]),int(todayStr[-1]),int(todayStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
				else:
					# print("Else Block")
					# print(t)
					start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
					end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
					# timeshiftB.append(str(t[0])+":"+str(t[1])+"-"+str(t[2])+":"+str(t[3]))
					start=tz.localize(start)
					end=tz.localize(end)
					data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
					if data:
						isQrData.append(1)
					else:
						isQrData.append(0)
			unAttendedShiftB.append(isQrData.count(0))
			AttendedShiftB.append(isQrData.count(1))
			isDataShiftB.append(isQrData)


	# isDataShiftG = []
	# AttendedShiftG = []
	# unAttendedShiftG = []
	# for qrid in uuidListG:
	# 	isQrData = []
	# 	for t in shiftG:
			
	# 		start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
	# 		end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
	# 		start=tz.localize(start)
	# 		end=tz.localize(end)
	# 		data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
	# 		if data:
	# 			isQrData.append(1)
	# 		else:
	# 			isQrData.append(0)
	# 	unAttendedShiftG.append(isQrData.count(0))
	# 	AttendedShiftG.append(isQrData.count(1))
	# 	isDataShiftG.append(isQrData)


	# print(AttendedShiftA,AttendedShiftB)





	maxAShiftA = max(AttendedShiftA)
	maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	# maxAShiftC = max(AttendedShiftC)
	# maxUShiftC = max(unAttendedShiftC)
	# maxAShiftG = max(AttendedShiftG)
	# maxUShiftG = max(unAttendedShiftG)
	
	# maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)
	# if siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels":
	# 	filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime.xlsx")
	# elif siteid=="Apollo Hospitals":
	# 	filename=os.path.join(module_dir,"Patrol Monitoring Report Format Two Shift.xlsx")
	# elif siteid=="Himadri Chemical Ltd.":
	# 	filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime1.xlsx")

	# else:
	# 	filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift.xlsx")

	if siteid!="Himadri Chemical Ltd.":
		filename=os.path.join(module_dir,"Patrol Monitoring Report Format Two Shift.xlsx")
	else:
		filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime1.xlsx")


	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	if siteid!="Apollo Hospitals":
		filename1 =siteid+" Post Monitoring Report.xlsx"
	else:
		filename1 =siteid+" Patrol Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	ws1 = wb2["Shift A"]
	ws2 = wb2["Shift B"]
	# ws3 = wb2["Shift C"]
	# ws4 = wb2["Shift G"]
	ws1['A4'] = siteid
	ws2['A4'] = siteid
	# ws3['A4'] = siteid
	# ws4['A4'] = siteid
	if siteid=="Apollo Hospitals":
		ws1['A3']="Patrol Monitoring System"
		ws2['A3']="Patrol Monitoring System"

	if siteid=="Ashwini Hospital":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"
	elif siteid=="The Methodist School" or siteid=="Ahura Mazda" or siteid=="SRG Exim SEIM00001" or siteid=="Hathway Digital HWCC01" or siteid=="ACC Limited SIL ACCL01" or siteid=="Aum Packaging APKG01":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 19:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 07:00 Hours"

	elif siteid=="KONE ELEVATOR INDIA":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 07:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 14:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 21:00 Hours"
	elif siteid=="Manaksia Limited MNKS01" or siteid=="Spencer Retail SPRL10":
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"


	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"


	rownum1=19
	thin = Side (border_style="thin", color="000000")

	for ele_ta in timeshiftA:
		ele_ta1=ele_ta.split("-")
		ws1.cell(row=rownum1,column=1).value=ele_ta1[0]
		ws1.cell(row=rownum1,column=2).value=ele_ta1[1]
		ws1.cell(row=rownum1,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row=rownum1,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum1=rownum1+1



	shiftAcol = 3
	shiftAi = 0
	for postA in place_list_siteidA:
		shiftArow = 19
		for data in isDataShiftA[shiftAi]:
			# ws1.cell(row = shiftArow, column = shiftAcol, value=post)
			# ws1.cell(row = shiftArow, column = shiftAcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws1.cell(row = shiftArow, column = shiftAcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws1.cell(row = shiftArow, column = shiftAcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				
			shiftArow += 1
		shiftAi += 1
		shiftAcol += 1


	rownum2=19
	for ele_tb in timeshiftB:
		ele_tb1=ele_tb.split("-")
		ws2.cell(row=rownum2,column=1).value=ele_tb1[0]
		ws2.cell(row=rownum2,column=2).value=ele_tb1[1]
		ws2.cell(row=rownum2,column=1).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row=rownum2,column=2).border=Border(top=thin, left=thin, right=thin, bottom=thin)
		rownum2=rownum2+1

	shiftBcol = 3
	shiftBi = 0
	for postB in place_list_siteidB:
		shiftBrow = 19
		for data in isDataShiftB[shiftBi]:
			# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
			# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
			if data:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			else:
				ws2.cell(row = shiftBrow, column = shiftBcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
				ws2.cell(row = shiftBrow, column = shiftBcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
			shiftBrow += 1
		shiftBi += 1
		shiftBcol += 1

	# shiftCcol = 3
	# shiftCi = 0
	# for postC in place_list_siteidC:
	# 	shiftCrow = 19
	# 	for data in isDataShiftC[shiftCi]:
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws3.cell(row = shiftCrow, column = shiftCcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		shiftCrow += 1
	# 	shiftCi += 1
	# 	shiftCcol += 1


	# shiftGcol = 3
	# shiftGi = 0
	# for postG in place_list_siteidG:
	# 	shiftGrow = 19
	# 	for data in isDataShiftG[shiftGi]:
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol, value=post)
	# 		# ws2.cell(row = shiftBrow, column = shiftBcol).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 		if data:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = '64c620', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		else:
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).fill = PatternFill(fgColor = 'fa0505', fill_type="solid")
	# 			ws4.cell(row = shiftGrow, column = shiftGcol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 		shiftGrow += 1
	# 	shiftGi += 1
	# 	shiftGcol += 1

	col = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws1.cell(row = 18, column = col, value=place_list_siteidA[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws1.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws1.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws1['C17'] = 'Post Name'

	ws1['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidA)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	col = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
		ws2.cell(row = 18, column = col, value=place_list_siteidB[post])
		
		# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		ws2.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
		# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

		# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		ws2.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		# ws1.column_dimensions[get_column_letter(col)].width = 25
		ws2.column_dimensions[get_column_letter(col)].width = 25
		# ws1.row_dimensions[18].height = 70
		ws2.row_dimensions[18].height = 70
		col += 1

	# ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	ws2.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	ws2['C17'] = 'Post Name'

	ws2['C17'].font = Font(bold=True)

	# ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	ws2.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	ws2.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	bordercol = 3
	for post in range(0, len(place_list_siteidB)):
		# ws1.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		ws2.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		bordercol += 1


	# col = 3
	# for post in range(0, len(place_list_siteidC)):
	# 	# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
	# 	ws3.cell(row = 18, column = col, value=place_list_siteidC[post])
		
	# 	# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 	ws3.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
	# 	# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws3.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

	# 	# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
	# 	ws3.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

	# 	# ws1.column_dimensions[get_column_letter(col)].width = 25
	# 	ws3.column_dimensions[get_column_letter(col)].width = 25
	# 	# ws1.row_dimensions[18].height = 70
	# 	ws3.row_dimensions[18].height = 70
	# 	col += 1

	# # ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	# ws3.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	# ws3['C17'] = 'Post Name'

	# ws3['C17'].font = Font(bold=True)

	# # ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws3.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# # ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# ws3.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	# bordercol = 3
	# for post in range(0, len(place_list_siteidC)):
	# 	# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws3.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	bordercol += 1


	# col = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 18, column = col, value=place_list_siteid[post])
	# 	ws4.cell(row = 18, column = col, value=place_list_siteidG[post])
		
	# 	# ws1.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# 	ws4.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		
	# 	# ws1.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)

	# 	# ws1.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
	# 	ws4.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

	# 	# ws1.column_dimensions[get_column_letter(col)].width = 25
	# 	ws4.column_dimensions[get_column_letter(col)].width = 25
	# 	# ws1.row_dimensions[18].height = 70
	# 	ws4.row_dimensions[18].height = 70
	# 	col += 1

	# # ws1.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)
	# ws4.merge_cells(start_row = 17, end_row=17, start_column=3, end_column = col-1)

	# ws4['C17'] = 'Post Name'

	# ws4['C17'].font = Font(bold=True)

	# # ws1.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# ws4.cell(row = 17, column = 3).fill = PatternFill(fgColor = 'D8D2D5', fill_type="solid")
	# # ws1.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
	# ws4.cell(row = 17, column = 3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	# bordercol = 3
	# for post in range(0, len(place_list_siteidG)):
	# 	# ws1.cell(row = 17, column = borderc).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	ws4.cell(row = 17, column = bordercol).border = Border(top=thin, left=thin, right=thin, bottom=thin)
	# 	bordercol += 1



	ws1['E8'] = place_list_siteidA[unAttendedShiftA.index(maxUShiftA)]
	ws1['E9'] = place_list_siteidA[AttendedShiftA.index(maxAShiftA)]
	ws1['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['E9'].alignment = Alignment(horizontal='center', vertical='center')
	# ws1.column_dimensions['F'].auto_size = True

	ws2['E8'] = place_list_siteidB[unAttendedShiftB.index(maxUShiftB)]
	ws2['E9'] = place_list_siteidB[AttendedShiftB.index(maxAShiftB)]
	ws2['E8'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws3['E8'] = place_list_siteidC[unAttendedShiftC.index(maxUShiftC)]
	# ws3['E9'] = place_list_siteidC[AttendedShiftC.index(maxAShiftC)]
	# ws3['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws3['E9'].alignment = Alignment(horizontal='center', vertical='center')

	# ws4['E8'] = place_list_siteidG[unAttendedShiftG.index(maxUShiftG)]
	# ws4['E9'] = place_list_siteidG[AttendedShiftG.index(maxAShiftG)]
	# ws4['E8'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['E9'].alignment = Alignment(horizontal='center', vertical='center')

	ws1['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws1['B18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['A18'].alignment = Alignment(horizontal='center', vertical='center')
	ws2['B18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws3['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws3['B18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['A18'].alignment = Alignment(horizontal='center', vertical='center')
	# ws4['B18'].alignment = Alignment(horizontal='center', vertical='center')

	# saving the destination excel file

	if siteid=="Himadri Chemical Ltd.":

		wb2.remove(wb2['Shift A'])

		ws3 = wb2["Timestamp"]

		time1=datetime.now()
		time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		time2_1=time2_1+timedelta(hours=8)
		time2_11=time2_1-timedelta(days=1)

		start_time1=tz.localize(time2_11)
		end_time1=tz.localize(time2_1)

		all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=siteid,time__range=(start_time1,end_time1)).values()))

		if len(all_path)!=0:

			all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
			all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

			df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')


			all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
			all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

			df_all=pd.merge(df_all1,all_emp,on='empid_id')

			df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

			# df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
			df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
			df_all=df_all.sort_values("time",ascending=True)
			df_all=df_all.reset_index()

			# print(df_all.columns)

			rownum=2

			for idx,rows in df_all.iterrows():
				ws3.cell(row = rownum, column = 1, value=rows['name'])
				ws3.cell(row = rownum, column = 2, value=rows['placeid_id'])
				ws3.cell(row = rownum, column = 3, value=rows['datetime1'])
				ws3.cell(row = rownum, column = 4, value=rows['empname'])

				ws3.cell(row = rownum, column = 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 2).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 3).border = Border(top=thin, left=thin, right=thin, bottom=thin)
				ws3.cell(row = rownum, column = 4).border = Border(top=thin, left=thin, right=thin, bottom=thin)

				rownum=rownum+1



			wb2.save(str(filename1))



		else:
			wb2.save(str(filename1))

	# elif siteid=="ALIPUR TERRACE":
	# 	wb2.remove(wb2['Shift A'])
	# 	wb2.save(str(filename1))


	else:
		wb2.save(str(filename1))



	if siteid=="Teloijan-Tea":
		sitenamemail="Poddar House(Teloijan-Tea)"
	elif siteid=="Thakur Prasad Sao":
		sitenamemail="Thakur Prasad Sao and Sons"
	else:
		sitenamemail=siteid


	try:
		# print("program is here")
		mail_list_path=os.path.join(module_dir,'mail_id_list.xlsx')
		df=pd.read_excel(mail_list_path, sheet_name='Sheet1', engine='openpyxl')

		# print(df)

		df_1=df.loc[df['siteid']==siteid]
		mail_list=df_1.iloc[:,1:].values.tolist()
		mail_list=[element for element in mail_list[0] if '@' in str(element)]
		send_to_list=mail_list
		#send_to_list=['shyamal@nis.co.in','aditya@achillesresolute.com','kanad@achillesresolute.com','suman@achillesresolute.com','bbsr.abhishek@nis.co.in','bbsr.pabitra@nis.co.in','santosh.prusti@amribbsr.in']
		# send_to_list=['kanad@achillesresolute.com','arnab@achillesresolute.com']#,'kanad@achillesresolute.com','suman@achillesresolute.com','aditya@achillesresolute.com']
		subject=f"{siteid}- Daily Post Monitoring Report"
		try:
			body=f"Dear Sir,\n\nPlease find attached the automated daily post monitoring report pertaining to {sitenamemail} for the date of {dt_1}.\n\nP.S: This is an autogenerated email. Please do not reply on this email.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		except:
			body=f"Dear Sir,\n\nPlease find attached the daily post monitoring report.\n\nThank you!\n\nKind Regards,\nPatrol Monitoring System\nNIS Management Limited"
		# html_message = render_to_string('table/table.html')
		# text_content=strip_tags(html_message)
		# message = get_template('table/table_1.html').render()
		# email=EmailMessage(subject,message,settings.EMAIL_HOST_USER, send_to_list)
		# email.content_subtype ="html"# Main content is now text/html

	# 	# email = EmailMultiAlternatives(subject,text_content,settings.EMAIL_HOST_USER, send_to_list)
	# 	# email.attach_alternative(html_message,"text/html")

		mail = EmailMessage(subject,body, settings.EMAIL_HOST_USER, send_to_list)
		mail.attach_file(filename1)
		mail.send()
		# print("done!")
		# return HttpResponse("Hello World", content_type="text/plain")
	except:
		send_to_list=['arnab@achillesresolute.com']
		body=body=f"Dear Sir,\n\nPlease find attached the automated monthly patrol performance report pertaining to {siteid} for the date of {dt_1}, is not delivered.Please check the matter.\n\nThank you!\n\nKind Regards,\nRemote Monitoring System\nNIS Management Limited"
		mail = EmailMessage(siteid+" | Post monitoring report failed",body, settings.EMAIL_HOST_USER, send_to_list)
		mail.send()


	return HttpResponse("Success")


def masterdashboard(request):

	authkey=request.GET.get("authkey","")
	empid=request.GET.get("empid","")
	branch=request.GET.get("branch","")
	auth_org="TjWnZr4u"

	# print(authkey,empid)

	df_site=pd.DataFrame(list(siteTbl.objects.all().values()))

	# print(df_site)
	df_site=df_site.loc[(df_site["name"]!="Test-site3")]
	df_site1=df_site.loc[(df_site["name"]!="Test-site2")]

	df_site1=df_site1.reset_index(drop=True)

	loclist=df_site1["location"].tolist()

	loc_list_final=[]

	for ele in loclist:
		if ele=="Guwahati":
			ele="Assam"
		else:
			ele=ele
		if ele not in loc_list_final:
			loc_list_final.append(ele)

	# df_grp=df_site1.groupby("location")
	# print(loc_list_final)

	if branch=="":
		loc_list_final1=loc_list_final
	else:
		loc_list_final1=[x for x in loc_list_final if x==branch]

	context={"data":loc_list_final1}


	# print(df_site1)

	return render(request,"master.html",context)


# def calculate_success_rate(qrcount,df,slotcount,type):


def postlist(request):
	branch=request.GET.get("branch","")


	if branch=="Assam":
		sitelist=list(siteTbl.objects.filter(Q(location=branch)|Q(location="Guwahati")).values_list('name',flat=True))

	else:
		sitelist=list(siteTbl.objects.filter(location=branch).values_list('name',flat=True))
		# print(sitelist)

	if branch=="Kolkata":

		sitelist.remove("Test-site2")
		sitelist.remove("Test-site3")
		sitelist.remove("NIS-Head Office")

	module_dir = os.path.dirname(__file__)

	mail_list_path=os.path.join(module_dir,'sitedetails.xlsx')
	df=pd.read_excel(mail_list_path, sheet_name='post', engine='openpyxl')
	# print(df)
	# # print(df)
	df1=df.loc[df["siteid"].isin(sitelist)]
	# df1.fillna("Not Applicable",inplace=True)
	df1=df1.reset_index(drop=True)
	# df1=df1.rename(columns={"Shift A freq":"A_freq","Shift B freq":"B_freq","Shift C freq":"C_freq","Shift G freq":"G_freq"})

	sitels=df1["siteid"].tolist()

	tz=timezone.get_current_timezone()

	time1=datetime.now()
	success_ls=[]
	success_ls_n=[]
	suc_old=[]
	suc_old_n=[]


	for ele in sitels:

		starthr=int(df1.loc[df1["siteid"]==ele,"A1"].tolist()[0])
		freq=int(df1.loc[df1["siteid"]==ele,"freq"].tolist()[0])
		nightonly=int(df1.loc[df1["siteid"]==ele,"night"].tolist()[0])
		dayonly=int(df1.loc[df1["siteid"]==ele,"day"].tolist()[0])

		startc=int(df1.loc[df1["siteid"]==ele,"C1"].tolist()[0])
		startb=int(df1.loc[df1["siteid"]==ele,"B1"].tolist()[0])
		satrtg=int(df1.loc[df1["siteid"]==ele,"G1"].tolist()[0])

		offseta=str(starthr)+"H"
		offsetb=str(startb)+"H"
		offsetc=str(startc)+"H"

		end_date1=time1-timedelta(days=1)
		end_date1=end_date1.date()
		start_date1=end_date1-timedelta(days=6)
		# start_date1=start_date1+timedelta(hours=6)

		# end_date1=tz.localize(end_date1)
		# start_date1=tz.localize(start_date1)
		endt=end_date1

		# print("This is week")

		# print(start_date1,end_date1)

		start_date2=start_date1-timedelta(days=7)
		end_date2=end_date1-timedelta(days=7)

		# print("This is last week")

		# print(start_date2,end_date2)

		dt1=datetime.strftime(start_date1,"%d %b")+"-"+datetime.strftime(endt,"%d %b")
		dt2=datetime.strftime(start_date2,"%d %b")+"-"+datetime.strftime(end_date2,"%d %b")





		df_success=pd.DataFrame(list(successTbl.objects.filter(siteid__name=ele,date1__range=(start_date1,end_date1)).values()))

		if len(df_success)==0:
			success_ls.append(0)
			success_ls_n.append(0)

		elif dayonly==1:
			suc_ovr=((round(df_success["succrate"].mean(),2)*2)-round(df_success["nightsuccrate"].mean(),2))
			success_ls.append(suc_ovr)
			success_ls_n.append(None)


		elif nightonly==1:
			if (round(df_success["succrate"].mean(),2)!=round(df_success["nightsuccrate"].mean(),2)):
				success_ls_n.append(round(df_success["nightsuccrate"].mean(),2))
				success_ls.append(round(df_success["nightsuccrate"].mean(),2))
			else:
				success_ls.append(round(df_success["succrate"].mean(),2))
				success_ls_n.append(round(df_success["nightsuccrate"].mean(),2))




		else:


			success_ls.append(round(df_success["succrate"].mean(),2))
			success_ls_n.append(round(df_success["nightsuccrate"].mean(),2))


		df_success_old=pd.DataFrame(list(successTbl.objects.filter(siteid__name=ele,date1__range=(start_date2,end_date2)).values()))


		if len(df_success_old)==0:
			suc_old.append(0)
			suc_old_n.append(0)

		elif nightonly==1:
			if (round(df_success_old["succrate"].mean(),2)!=round(df_success_old["nightsuccrate"].mean(),2)):
				suc_old.append(round(df_success_old["nightsuccrate"].mean(),2))
				suc_old_n.append(round(df_success_old["nightsuccrate"].mean(),2))
			else:
				suc_old.append(round(df_success_old["succrate"].mean(),2))
				suc_old_n.append(round(df_success_old["nightsuccrate"].mean(),2))


		else:

			suc_old.append(round(df_success_old["succrate"].mean(),2))
			suc_old_n.append(round(df_success_old["nightsuccrate"].mean(),2))








	df1["success"]=success_ls
	df1["successn"]=success_ls_n


	# df1.to_csv("test-assam.csv")


	
	df1["oldsucc"]=suc_old
	df1["oldnight"]=suc_old_n
	df1["overdelta"]=df1["success"]-df1["oldsucc"]
	df1["nightd"]=df1["successn"]-df1["oldnight"]

	df1["color1"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["overdelta"]]
	df1["color2"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["nightd"]]

	df1["arrow1"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["overdelta"]]
	df1["arrow2"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["nightd"]]





	ovr_success=round(df1["success"].mean(),2)
	ovr_old=round(df1["oldsucc"].mean(),2)

	delta_ovr=round(ovr_success-ovr_old,2)

	if delta_ovr<0:
		cl1="fa fa-arrow-down"
		col1="color:#ff0000"
	else:
		cl1="fa fa-arrow-up"
		col1="color:#00ff00"


	ovr_night=round(df1["successn"].mean(),2)
	ovr_night_old=round(df1["oldnight"].mean(),2)

	delta_night=round(ovr_night-ovr_night_old,2)

	if delta_night<0:
		cl2="fa fa-arrow-down"
		col2="color:#ff0000"
	else:
		cl2="fa fa-arrow-up"
		col2="color:#00ff00"





	








	json_records = df1.reset_index().to_json(orient ='records')
	data = []
	data = json.loads(json_records)

	context={"branch":branch,"data":data,"ovr_s":ovr_success,"ovr_old":ovr_old,"ovr_n":ovr_night,"ovr_no":ovr_night_old,"del_ovr":delta_ovr,"del_n":delta_night,
	"cl1":cl1,"col1":col1,"cl2":cl2,"col2":col2,"dt1":dt1,"dt2":dt2,"totsite":len(df1)}

	#return HttpResponse("Success")

	return render(request,"postlist.html",context)








def calculate_score(request):
	branch=request.GET.get("branch","")

	days=request.GET.get("days","")
	date1=request.GET.get("date1","")

	print(branch)



	if branch=="Assam":
		sitelist=list(siteTbl.objects.filter(Q(location=branch)|Q(location="Guwahati")).values_list('name',flat=True))

	else:
		sitelist=list(siteTbl.objects.filter(location=branch).values_list('name',flat=True))
		print(sitelist)


	if branch=="Kolkata":

		sitelist.remove("Test-site2")
		sitelist.remove("Test-site3")
		sitelist.remove("NIS-Head Office")
	# print(sitelist)








	module_dir = os.path.dirname(__file__)

	mail_list_path=os.path.join(module_dir,'sitedetails1.xlsx')
	df=pd.read_excel(mail_list_path, sheet_name='post', engine='openpyxl')
	# print(df)
	# # print(df)
	df1=df.loc[df["siteid"].isin(sitelist)]
	# df1.fillna("Not Applicable",inplace=True)
	df1=df1.reset_index(drop=True)
	# df1=df1.rename(columns={"Shift A freq":"A_freq","Shift B freq":"B_freq","Shift C freq":"C_freq","Shift G freq":"G_freq"})

	sitels=df1["siteid"].tolist()

	tz=timezone.get_current_timezone()

	time1=datetime.now()


	if days=="":
		days_data=1
		# end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
		# end_date1=end_date1+timedelta(hours=starthr)
		# start_date1=end_date1-timedelta(days=7)
	else:
		days=int(float(days))
		days_data=days

	datels=[]
	site_ls=[]
	succ_ls=[]
	succ_n_ls=[]
	print(sitels)



	for ele in sitels:
		all_beacons_names=list(uuidTbl.objects.filter(placeid__siteid__name=ele,uuidtype="QR").values_list("uuid",flat=True))
		totqr=len(all_beacons_names)

		starthr=int(df1.loc[df1["siteid"]==ele,"A1"].tolist()[0])
		freq=int(df1.loc[df1["siteid"]==ele,"freq"].tolist()[0])
		nightonly=int(df1.loc[df1["siteid"]==ele,"night"].tolist()[0])
		dayonly=int(df1.loc[df1["siteid"]==ele,"day"].tolist()[0])

		startc=int(df1.loc[df1["siteid"]==ele,"C1"].tolist()[0])
		startb=int(df1.loc[df1["siteid"]==ele,"B1"].tolist()[0])
		satrtg=int(df1.loc[df1["siteid"]==ele,"G1"].tolist()[0])

		offseta=str(starthr)+"H"
		offsetb=str(startb)+"H"
		offsetc=str(startc)+"H"

		freq_sample=str(float(freq))+"min"


		if date1=="":

			end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			end_date1=end_date1+timedelta(hours=starthr)
			start_date1=end_date1-timedelta(days=days_data)
			# start_date1=start_date1+timedelta(hours=6)

			end_date1=tz.localize(end_date1)
			start_date1=tz.localize(start_date1)
			endt=end_date1-timedelta(days=1)

		else:

			date2=datetime.strptime(date1, '%Y-%m-%d')
			date2=date2+timedelta(days=1)

			# end_date1=time1-timedelta(hours=date2.hour,minutes=date2.minute)
			end_date1=date2+timedelta(hours=starthr)
			start_date1=end_date1-timedelta(days=days_data)
			# start_date1=start_date1+timedelta(hours=6)

			end_date1=tz.localize(end_date1)
			start_date1=tz.localize(start_date1)
			endt=end_date1-timedelta(days=1)

			print("This is date", start_date1,end_date1)




		# tot_slot=24


		if nightonly==1:
			# end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			# end_date1=end_date1+timedelta(hours=starthr)
			# start_date1=end_date1-timedelta(days=days_data)
			# # start_date1=start_date1+timedelta(hours=6)

			# end_date1=tz.localize(end_date1)
			# start_date1=tz.localize(start_date1)
			# endt=end_date1-timedelta(days=1)

			dt1=datetime.strftime(endt,"%d %b %Y")




			# dt1=datetime.strftime(start_date1,"%d %b")+"-"+datetime.strftime(endt,"%d %b")

			# print(start_date1,end_date1)

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name=ele,uuid__uuidtype="QR",time__range=(start_date1,end_date1)).values()))
			# df_old=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date_old,end_date_old)).values()))

			if len(df_all)==0:
				success_ovr=0
				success_n=0
				success_g=0

			else:
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour


				if startc==0:
					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr

					print("Total Hour is", total_hour)

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/freq)

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ovr=ovr_succ
					success_n=ovr_succ
					success_g=0

				else:

					#latest 7 days start


					df_all2=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
					# print(startb,starthr)
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

					total_hour=(24-startc)+starthr



					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/freq)


					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ovr=ovr_succ
					success_n=ovr_succ
					success_g=0
		elif dayonly==1:
			# end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			# end_date1=end_date1+timedelta(hours=starthr)
			# start_date1=end_date1-timedelta(days=days_data)
			# # start_date1=start_date1+timedelta(hours=6)

			# end_date1=tz.localize(end_date1)
			# start_date1=tz.localize(start_date1)
			# endt=end_date1-timedelta(days=1)




			dt1=datetime.strftime(endt,"%d %b %Y")

			# print(start_date1,end_date1)

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name=ele,uuid__uuidtype="QR",time__range=(start_date1,end_date1)).values()))

			if len(df_all)==0:
				success_ovr=0
				success_n=None
				success_g=0

			else:
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour


				if startc==0:
					df_all2=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr

					print("Total Hour is", total_hour)

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/freq)

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ovr=ovr_succ
					success_n=None
					success_g=0

				else:

					#latest 7 days start


					df_all2=df_all.loc[(df_all["hour"]>=starta)&(df_all["hour"]<starthb)]
					# print(startb,starthr)
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

					total_hour=(24-startc)+starthr



					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/freq)


					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ovr=ovr_succ
					success_n=None
					success_g=0






		else:
			# end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			# end_date1=end_date1+timedelta(hours=starthr)
			# start_date1=end_date1-timedelta(days=days_data)
			# # start_date1=start_date1+timedelta(hours=6)

			# end_date1=tz.localize(end_date1)
			# start_date1=tz.localize(start_date1)
			# endt=end_date1-timedelta(days=1)




			dt1=datetime.strftime(endt,"%d %b %Y")

			# print(start_date1,end_date1)

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name=ele,uuid__uuidtype="QR",time__range=(start_date1,end_date1)).values()))
			# df_old=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date_old,end_date_old)).values()))

			if len(df_all)==0:
				success_ovr=0
				success_n=0
				success_g=0

			else:
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				if startc==0 and freq!=45:
					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)
					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(startb-starthr)

					
					df_a=df_all1[["time","uuid_id"]]
					df_a1=df_a.resample(freq_sample).nunique()
					# print(type(df_a1))
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=days_data*total_hour*totqr_a*(60/freq)

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/freq)

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=(succrate_a+succrate_b)/2

					success_ovr=ovr_succ
					success_n=succrate_b
					success_g=0
				elif startc==0 and freq==45:
					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)
					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(startb-starthr)

					
					df_a=df_all1[["time","uuid_id"]]
					df_a1=df_a.resample('60min').nunique()
					# print(type(df_a1))
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=days_data*total_hour*totqr_a*(60/60)

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample('30min').nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=days_data*total_hour*totqr_b*(60/30)

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=(succrate_a+succrate_b)/2

					success_ovr=ovr_succ
					success_n=succrate_b
					success_g=0
				else:
					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)

					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
					df_all2=df_all2.reset_index(drop=True)

					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
					df_all3=df_all3.reset_index(drop=True)

					df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
					totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

					total_hour_a=(startb-starthr)
					total_hour_b=(startc-startb)
					total_hour_c=(24-startc)+starthr

					print(total_hour_a,total_hour_b,total_hour_c,freq_sample)

					
					df_a=df_all1[["time","uuid_id"]]
					df_a1=df_a.resample(freq_sample).nunique()
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=total_hour_a*totqr_a*(60/freq)

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample(freq_sample).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=total_hour_b*totqr_b*(60/freq)

					succrate_b=(tot_succ_b/tot_org_b)*100

					df_c=df_all3[["time","uuid_id"]]
					df_c1=df_c.resample(freq_sample).nunique()
					df_c1.loc[df_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_c=df_c1["uuid_id"].sum()

					tot_org_c=total_hour_c*totqr_c*(60/freq)

					succrate_c=(tot_succ_c/tot_org_c)*100




					ovr_succ=(succrate_a+succrate_b+succrate_c)/3


					success_ovr=ovr_succ
					success_n=succrate_c
					success_g=0

		succ_ls.append(success_ovr)
		succ_n_ls.append(success_n)
		site_ls.append(ele)
		datels.append(dt1)


	df_main=pd.DataFrame({"site":site_ls,"succ":succ_ls,"night":succ_n_ls,"date":datels})
	df_main["date"]=pd.to_datetime(df_main["date"],format="%d %b %Y")

	for index,rows in df_main.iterrows():

		siteData=siteTbl.objects.get(siteid=rows["site"])
		x=successTbl.objects.filter(siteid=siteData,date1=rows["date"]).exists()
		if x==False:
			abc=successTbl(siteid=siteData,date1=rows["date"],succrate=rows["succ"],nightsuccrate=rows["night"]).save()








	# df_main.to_csv("datatest.csv")


	return HttpResponse("Success")

















def postlist_data(request):

	branch=request.GET.get("branch","")

	if branch=="Assam":
		sitelist=list(siteTbl.objects.filter(Q(location=branch)|Q(location="Guwahati")).values_list('name',flat=True))

	else:
		sitelist=list(siteTbl.objects.filter(location=branch).values_list('name',flat=True))




	if branch=="Kolkata":

		sitelist.remove("Test-site2")
		sitelist.remove("Test-site3")
		sitelist.remove("NIS-Head Office")
	# print(sitelist)






	module_dir = os.path.dirname(__file__)

	mail_list_path=os.path.join(module_dir,'sitedetails.xlsx')
	df=pd.read_excel(mail_list_path, sheet_name='post', engine='openpyxl')
	# print(df)
	df1=df.loc[df["siteid"].isin(sitelist)]
	# df1.fillna("Not Applicable",inplace=True)
	df1=df1.reset_index(drop=True)
	# df1=df1.rename(columns={"Shift A freq":"A_freq","Shift B freq":"B_freq","Shift C freq":"C_freq","Shift G freq":"G_freq"})

	sitels=df1["siteid"].tolist()

	tz=timezone.get_current_timezone()

	time1=datetime.now()

	
	# print(len(df4))


	success_ls=[]
	success_ls_n=[]
	suc_old=[]
	suc_old_n=[]









	for ele in sitels:
		all_beacons_names=list(uuidTbl.objects.filter(placeid__siteid__name=ele,uuidtype="QR").values_list("uuid",flat=True))
		totqr=len(all_beacons_names)

		starthr=int(df1.loc[df1["siteid"]==ele,"A1"].tolist()[0])
		freq=int(df1.loc[df1["siteid"]==ele,"freq"].tolist()[0])
		nightonly=int(df1.loc[df1["siteid"]==ele,"night"].tolist()[0])
		startc=int(df1.loc[df1["siteid"]==ele,"C1"].tolist()[0])
		startb=int(df1.loc[df1["siteid"]==ele,"B1"].tolist()[0])

		# tot_slot=24


		if nightonly==1:
			end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			end_date1=end_date1+timedelta(hours=starthr)
			start_date1=end_date1-timedelta(days=7)
			# start_date1=start_date1+timedelta(hours=6)

			end_date1=tz.localize(end_date1)
			start_date1=tz.localize(start_date1)
			endt=end_date1-timedelta(days=1)

			start_date_old=start_date1-timedelta(days=7)
			end_date_old=end_date1-timedelta(days=7)
			endt1=end_date_old-timedelta(days=1)

			dt1=datetime.strftime(start_date1,"%d %b")+"-"+datetime.strftime(endt,"%d %b")
			dt2=datetime.strftime(start_date_old,"%d %b")+"-"+datetime.strftime(endt1,"%d %b")

			# print(start_date1,end_date1)

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date1,end_date1)).values()))
			df_old=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date_old,end_date_old)).values()))

			if len(df_all)==0 and len(df_old)==0:
				success_ls.append(0)
				success_ls_n.append(0)
				suc_old.append(0)
				suc_old_n.append(0)

			elif len(df_old)==0 and len(df_all)!=0:
				suc_old.append(0)
				suc_old_n.append(0)
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				if freq==30:
					if startc==0:
						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)
						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(24-startb)+starthr

						print("Total Hour is", total_hour)

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=succrate_b

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

				else:

					#latest 7 days start


					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					# print(startb,starthr)
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr



					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample('1H').nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=7*total_hour*totqr_b


					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ls_n.append(round(succrate_b,2))
					success_ls.append(round(ovr_succ,2))
					#Latest 7 days end

			elif len(df_old)!=0 and len(df_all)==0:

				success_ls_n.append(0)
				success_ls.append(0)
				df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_old["hour"]=df_old["time"].dt.hour

				if freq==30:
					if startc==0:
						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(24-startb)+starthr

						print("Total Hour is", total_hour)

						df_b=df_old2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=succrate_b

						suc_old.append(round(succrate_b,2))
						suc_old_n.append(round(ovr_succ,2))

				else:

					#latest 7 days start


					df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
					# print(startb,starthr)
					df_old2=df_old2.reset_index(drop=True)
					df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr



					df_b=df_old2[["time","uuid_id"]]
					df_b1=df_b.resample('1H').nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=7*total_hour*totqr_b


					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					suc_old.append(round(succrate_b,2))
					suc_old_n.append(round(ovr_succ,2))


			else:

				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_old["hour"]=df_old["time"].dt.hour

				if freq==30:
					if startc==0:
						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)
						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(24-startb)+starthr

						print("Total Hour is", total_hour)

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=succrate_b

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))


						#Old 7 days start
						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						# print(startb,starthr)
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(24-startb)+starthr



						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b
						# print("total slot",total_org_b)

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=succrate_b

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))
				else:

					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					print(startb,starthr)
					df_all2=df_all2.reset_index(drop=True)
					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr

					# print("Total Hour is", total_hour)

					df_b=df_all2[["time","uuid_id"]]
					df_b1=df_b.resample('1H').nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=7*total_hour*totqr_b
					# print("total slot",total_org_b)

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					success_ls_n.append(round(succrate_b,2))
					success_ls.append(round(ovr_succ,2))

					#Old 7 days start
					df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
					# print(startb,starthr)
					df_old2=df_old2.reset_index(drop=True)
					df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

					# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(24-startb)+starthr



					df_old_b=df_old2[["time","uuid_id"]]
					df_old_b1=df_old_b.resample('1H').nunique()
					df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_old_b1["uuid_id"].sum()

					tot_org_b=7*total_hour*totqr_b

					succrate_b=(tot_succ_b/tot_org_b)*100

					ovr_succ=succrate_b

					suc_old_n.append(round(succrate_b,2))
					suc_old.append(round(ovr_succ,2))




				# print(df3)

		else:
			end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			end_date1=end_date1+timedelta(hours=starthr)
			start_date1=end_date1-timedelta(days=7)
			# start_date1=start_date1+timedelta(hours=6)

			end_date1=tz.localize(end_date1)
			start_date1=tz.localize(start_date1)

			# print(start_date1,end_date1)



			start_date_old=start_date1-timedelta(days=7)
			end_date_old=end_date1-timedelta(days=7)
			# print(start_date_old,end_date_old)

			endt=end_date1-timedelta(days=1)
			endt1=end_date_old-timedelta(days=1)

			dt1=datetime.strftime(start_date1,"%d %b")+"-"+datetime.strftime(endt,"%d %b")
			dt2=datetime.strftime(start_date_old,"%d %b")+"-"+datetime.strftime(endt1,"%d %b")

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date1,end_date1)).values()))
			df_old=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR",time__range=(start_date_old,end_date_old)).values()))

			if len(df_all)==0 and len(df_old)==0:
				success_ls.append(0)
				success_ls_n.append(0)
				suc_old.append(0)
				suc_old_n.append(0)
			elif len(df_all)!=0 and len(df_old)==0:
				suc_old.append(0)
				suc_old_n.append(0)

				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				if freq==30:

					if startc==0:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)
						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)
						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('30min').nunique()
						# print(type(df_a1))
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))
					else:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
						df_all3=df_all3.reset_index(drop=True)

						df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('30min').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_c=df_all3[["time","uuid_id"]]
						df_c1=df_c.resample('30min').nunique()
						df_c1.loc[df_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c*2

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						success_ls_n.append(round(succrate_c,2))
						success_ls.append(round(ovr_succ,2))




				elif freq==45:
					if startc==0:
						df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a	#every 1 hour

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						df_b1=df_b.resample('30min').nunique()
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2	#every 1 hour

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

				elif freq==60:
					if startc==0:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('1H').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

					else:

						df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
						df_all3=df_all3.reset_index(drop=True)

						df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('1H').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_c=df_all3[["time","uuid_id"]]
						df_c1=df_c.resample('1H').nunique()
						tot_succ_c=df_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						success_ls_n.append(round(succrate_c,2))
						success_ls.append(round(ovr_succ,2))



			elif len(df_all)>0 and len(df_old)>0:



				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_old["hour"]=df_old["time"].dt.hour
				# df_all[""]

					# df3=df_all.set_index('time').groupby('uuid_id').resample('1D')['value'].ffill()

				if freq==30:

					if startc==0:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)
						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)
						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('30min').nunique()
						# print(type(df_a1))
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

						#old 7 days start
						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)
						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						# totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('30min').nunique()
						# print(type(df_old_a1))
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))
					else:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
						df_all3=df_all3.reset_index(drop=True)

						df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('30min').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_c=df_all3[["time","uuid_id"]]
						df_c1=df_c.resample('30min').nunique()
						df_c1.loc[df_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c*2

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						success_ls_n.append(round(succrate_c,2))
						success_ls.append(round(ovr_succ,2))


						#old 7 days start
						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)&(df_old["hour"]<startc)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						df_old3=df_old.loc[(df_old["hour"]>=startc)|(df_old["hour"]<starthr)]
						df_old3=df_old3.reset_index(drop=True)

						df_old3=df_old3.set_index(pd.to_datetime(df_old3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('30min').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_old_c=df_old3[["time","uuid_id"]]
						df_old_c1=df_old_c.resample('30min').nunique()
						df_old_c1.loc[df_old_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_old_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c*2

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						suc_old_n.append(round(succrate_c,2))
						suc_old.append(round(ovr_succ,2))




				elif freq==45:
					if startc==0:
						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a	#every 1 hour

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('30min').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2	#every 30 minutes

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

						#old 7 days
						# df_old=df_old.set_index(pd.to_datetime(df_old['time']))

						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a	#every 1 hour

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2	#every 30 minutes

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))

				elif freq==60:
					if startc==0:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('1H').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))


						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)
						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						# totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						# print(type(df_old_a1))
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('1H').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))
					else:

						df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
						df_all1=df_all1.reset_index(drop=True)

						df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

						df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
						df_all2=df_all2.reset_index(drop=True)

						df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

						df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
						df_all3=df_all3.reset_index(drop=True)

						df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_a=df_all1[["time","uuid_id"]]
						df_a1=df_a.resample('1H').nunique()
						df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_b=df_all2[["time","uuid_id"]]
						df_b1=df_b.resample('1H').nunique()
						df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_c=df_all3[["time","uuid_id"]]
						df_c1=df_c.resample('1H').nunique()
						tot_succ_c=df_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						success_ls_n.append(round(succrate_c,2))
						success_ls.append(round(ovr_succ,2))


						#old 7 days start


						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)&(df_old["hour"]<startc)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						df_old3=df_old.loc[(df_old["hour"]>=startc)|(df_old["hour"]<starthr)]
						df_old3=df_old3.reset_index(drop=True)

						df_old3=df_old3.set_index(pd.to_datetime(df_old3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('1H').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_old_c=df_old3[["time","uuid_id"]]
						df_old_c1=df_old_c.resample('1H').nunique()
						df_old_c1.loc[df_old_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_old_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c

						print(total_hour_c,totqr_c,ele)

						# print(tot_org_c)

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						suc_old_n.append(round(succrate_c,2))
						suc_old.append(round(ovr_succ,2))

			else:
				# df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				# df_all["hour"]=df_all["time"].dt.hour

				success_ls_n.append(0)
				success_ls.append(0)

				df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_old["hour"]=df_old["time"].dt.hour
				# df_all[""]

					# df3=df_all.set_index('time').groupby('uuid_id').resample('1D')['value'].ffill()

				if freq==30:

					if startc==0:

						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						

						#old 7 days start
						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)
						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						# totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('30min').nunique()
						# print(type(df_old_a1))
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))
					else:

						


						#old 7 days start
						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)&(df_old["hour"]<startc)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						df_old3=df_old.loc[(df_old["hour"]>=startc)|(df_old["hour"]<starthr)]
						df_old3=df_old3.reset_index(drop=True)

						df_old3=df_old3.set_index(pd.to_datetime(df_old3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('30min').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a*2

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b*2

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_old_c=df_old3[["time","uuid_id"]]
						df_old_c1=df_old_c.resample('30min').nunique()
						df_old_c1.loc[df_old_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_old_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c*2

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						suc_old_n.append(round(succrate_c,2))
						suc_old.append(round(ovr_succ,2))




				elif freq==45:
					if startc==0:
						# df_all=df_all.set_index(pd.to_datetime(df_all['time']))

						

						#old 7 days
						# df_old=df_old.set_index(pd.to_datetime(df_old['time']))

						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a	#every 1 hour

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('30min').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b*2	#every 30 minutes

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))

				elif freq==60:
					if startc==0:

						


						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)
						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
						df_old2=df_old2.reset_index(drop=True)
						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						# totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						# totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

						total_hour=(startb-starthr)

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						# print(type(df_old_a1))
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('1H').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))
					else:

						

						#old 7 days start


						df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
						df_old1=df_old1.reset_index(drop=True)

						df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

						df_old2=df_old.loc[(df_old["hour"]>=startb)&(df_old["hour"]<startc)]
						df_old2=df_old2.reset_index(drop=True)

						df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

						df_old3=df_old.loc[(df_old["hour"]>=startc)|(df_old["hour"]<starthr)]
						df_old3=df_old3.reset_index(drop=True)

						df_old3=df_old3.set_index(pd.to_datetime(df_old3['time']))

						totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
						totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
						totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

						total_hour_a=(startb-starthr)
						total_hour_b=(startc-startb)
						total_hour_c=(24-startc)+starthr

						
						df_old_a=df_old1[["time","uuid_id"]]
						df_old_a1=df_old_a.resample('1H').nunique()
						df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
						tot_succ_a=df_old_a1["uuid_id"].sum()

						tot_org_a=7*total_hour_a*totqr_a

						succrate_a=(tot_succ_a/tot_org_a)*100

						df_old_b=df_old2[["time","uuid_id"]]
						df_old_b1=df_old_b.resample('1H').nunique()
						df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
						tot_succ_b=df_old_b1["uuid_id"].sum()

						tot_org_b=7*total_hour_b*totqr_b

						succrate_b=(tot_succ_b/tot_org_b)*100

						df_old_c=df_old3[["time","uuid_id"]]
						df_old_c1=df_old_c.resample('1H').nunique()
						df_old_c1.loc[df_old_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
						tot_succ_c=df_old_c1["uuid_id"].sum()

						tot_org_c=7*total_hour_c*totqr_c

						print(total_hour_c,totqr_c,ele)

						# print(tot_org_c)

						succrate_c=(tot_succ_c/tot_org_c)*100




						ovr_succ=(succrate_a+succrate_b+succrate_c)/3



						suc_old_n.append(round(succrate_c,2))
						suc_old.append(round(ovr_succ,2))






	# 	starthr=int(starthr)

		# print(starthr)

		# time1=datetime.now()

		# end_date1=time1-timedelta(days=1,hours=time1.hour,minutes=time1.minute)

		# end_date1=tz.localize(end_date1)

		# date_list=[]
		# succ_list=[]


	df1["success"]=success_ls
	df1["successn"]=success_ls_n
	df1["oldsucc"]=suc_old
	df1["oldnight"]=suc_old_n
	df1["overdelta"]=df1["success"]-df1["oldsucc"]
	df1["nightd"]=df1["successn"]-df1["oldnight"]

	df1["color1"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["overdelta"]]
	df1["color2"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["nightd"]]

	df1["arrow1"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["overdelta"]]
	df1["arrow2"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["nightd"]]





	ovr_success=round(df1["success"].mean(),2)
	ovr_old=round(df1["oldsucc"].mean(),2)

	delta_ovr=round(ovr_success-ovr_old,2)

	if delta_ovr<0:
		cl1="fa fa-arrow-down"
		col1="color:#ff0000"
	else:
		cl1="fa fa-arrow-up"
		col1="color:#00ff00"


	ovr_night=round(df1["successn"].mean(),2)
	ovr_night_old=round(df1["oldnight"].mean(),2)

	delta_night=round(ovr_night-ovr_night_old,2)

	if delta_night<0:
		cl2="fa fa-arrow-down"
		col2="color:#ff0000"
	else:
		cl2="fa fa-arrow-up"
		col2="color:#00ff00"


	








	json_records = df1.reset_index().to_json(orient ='records')
	data = []
	data = json.loads(json_records)

	context={"branch":branch,"data":data,"ovr_s":ovr_success,"ovr_old":ovr_old,"ovr_n":ovr_night,"ovr_no":ovr_night_old,"del_ovr":delta_ovr,"del_n":delta_night,
	"cl1":cl1,"col1":col1,"cl2":cl2,"col2":col2,"dt1":dt1,"dt2":dt2,"totsite":len(df1)}


	# shift_timeA=[]
	# shift_timeB=[]
	# shift_timeC=[]
	# df1["A_freq"]=df1["Shift A freq"].astype
	# df1["shift_timeA"]=df1["Shift A"]+"\n"+df1["Shift A freq"].astype(str)+" minutes"
	# print(df1)

	# return HttpResponse("Success")

	return render(request,"postlist.html",context)

# @csrf_exempt
def patrolist(request):
	branch=request.GET.get("branch","")

	if branch=="Assam":
		sitelist=list(siteTbl.objects.filter(Q(location=branch)|Q(location="Guwahati")).values_list('name',flat=True))

	else:
		sitelist=list(siteTbl.objects.filter(location=branch).values_list('name',flat=True))




	if branch=="Kolkata":

		sitelist.remove("Test-site2")
		sitelist.remove("Test-site3")
		sitelist.remove("NIS-Head Office")
	# print(sitelist)




	module_dir = os.path.dirname(__file__)

	mail_list_path=os.path.join(module_dir,'sitedetails.xlsx')
	df=pd.read_excel(mail_list_path, sheet_name='patrol', engine='openpyxl')
	# print(df)
	if branch=="Kolkata":
		df1=df.loc[(df["siteid"].isin(sitelist))|(df["siteid"]=="AMRI-Dhakuria")|(df["siteid"]=="Utalika-Residence")]
	elif branch=="Bhubaneswar":
		df1=df.loc[(df["siteid"].isin(sitelist))|(df["siteid"]=="AMRI-Bhubaneswar")]
	else:
		df1=df.loc[(df["siteid"].isin(sitelist))]


	# df1.fillna("Not Applicable",inplace=True)
	df1=df1.reset_index(drop=True)
	df1=df1.rename(columns={"Shift A freq":"A_freq","Shift B freq":"B_freq","Shift C freq":"C_freq","Shift G freq":"G_freq"})


	sitels=df1["siteid"].tolist()

	success_ls=[]
	success_ls_n=[]
	suc_old_n=[]
	suc_old=[]

	time1=datetime.now()
	tz=timezone.get_current_timezone()
	strt1=time1-timedelta(days=7)
	end1=time1-timedelta(days=1)
	strt1_=strt1-timedelta(days=7)
	end1_=end1-timedelta(days=7)


	dt1=datetime.strftime(strt1,"%d %b")+"-"+datetime.strftime(end1,"%d %b")
	dt2=datetime.strftime(strt1_,"%d %b")+"-"+datetime.strftime(end1_,"%d %b")

	for ele in sitels:
		if (ele=="AMRI-Dhakuria") or (ele=="Utalika-Residence"):
			url="http://188.166.10.173:7878/lastsuccess?branch=Kolkata"
			try:
				a=requests.get(url)
				print(a)
				data=json.loads(a.text)
			except Exception as e:
				print(e)
			
			if ele=="AMRI-Dhakuria":
				success_ls.append(data["AMRI-Dhakuria-1"])
				success_ls_n.append(data["AMRI-Dhakuria-n1"])
				suc_old.append(data["AMRI-Dhakuria-2"])
				suc_old_n.append(data["AMRI-Dhakuria-n2"])
			elif ele=="Utalika-Residence":
				success_ls.append(data["Utalika-Residence-n1"])
				success_ls_n.append(data["Utalika-Residence-n1"])
				suc_old.append(data["Utalika-Residence-n2"])
				suc_old_n.append(data["Utalika-Residence-n2"])
			else:
				success_ls.append(data["Utalika-Residence-n1"])
				success_ls_n.append(data["Utalika-Residence-n2"])
				suc_old.append(data["Utalika-Residence-n1"])
				suc_old_n.append(data["Utalika-Residence-n2"])
		elif ele=="AMRI-Bhubaneswar":
			url="http://188.166.10.173:7878/lastsuccess?branch=Bhubaneswar"
			try:
				a=requests.get(url)
				print(a)
				data=json.loads(a.text)
			except Exception as e:
				print(e)

			success_ls.append(data["AMRI-Bhubaneswar-1"])
			success_ls_n.append(data["AMRI-Bhubaneswar-n1"])
			suc_old.append(data["AMRI-Bhubaneswar-2"])
			suc_old_n.append(data["AMRI-Bhubaneswar-n2"])

		else:
			print("program is here")

			starthr=int(df1.loc[df1["siteid"]==ele,"A1"].tolist()[0])
			freq=int(df1.loc[df1["siteid"]==ele,"freq"].tolist()[0])
			nightonly=int(df1.loc[df1["siteid"]==ele,"night"].tolist()[0])
			startc=int(df1.loc[df1["siteid"]==ele,"C1"].tolist()[0])
			startb=int(df1.loc[df1["siteid"]==ele,"B1"].tolist()[0])

			end_date1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
			end_date1=end_date1+timedelta(hours=starthr)
			start_date1=end_date1-timedelta(days=7)
			# start_date1=start_date1+timedelta(hours=6)

			

			end_date1=tz.localize(end_date1)
			start_date1=tz.localize(start_date1)

			start_date_old=start_date1-timedelta(days=7)
			end_date_old=end_date1-timedelta(days=7)

			print(start_date1,end_date1)

			print(start_date_old,end_date_old)

			endt=end_date1-timedelta(days=1)
			endt1=end_date_old-timedelta(days=1)

			dt1=datetime.strftime(start_date1,"%d %b")+"-"+datetime.strftime(endt,"%d %b")
			dt2=datetime.strftime(start_date_old,"%d %b")+"-"+datetime.strftime(endt1,"%d %b")

			df_all=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR PATROL",time__range=(start_date1,end_date1)).values()))
			df_old=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=ele,uuid__uuidtype="QR PATROL",time__range=(start_date_old,end_date_old)).values()))

			if len(df_all)==0 and len(df_old)==0:
				success_ls.append(0)
				success_ls_n.append(0)
				suc_old.append(0)
				suc_old_n.append(0)

			elif len(df_all)!=0 and len(df_old)==0:
				suc_old.append(0)
				suc_old_n.append(0)
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				# df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				# df_old["hour"]=df_old["time"].dt.hour
				if startc==0:
					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)

					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					df_all2=df_all2.reset_index(drop=True)

					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(startb-starthr)

					
					df_a=df_all1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_a1=df_a.resample(freqhr).nunique()
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_b1=df_b.resample(freqhr).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=(7*total_hour*totqr_b)/freq

					succrate_b=(tot_succ_b/tot_org_b)*100

					if ele=="Supreme Industries" or ele=="Supreme Industries SPRM01" or ele=="YORK PRINT YPRT03" or ele=="YORK PRINT YPRT04":
						ovr_succ=succrate_a
						success_ls.append(round(ovr_succ,2))
						success_ls_n.append(0)
					else:



						ovr_succ=(succrate_a+succrate_b)/2



						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))


					
				else:

					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)

					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
					df_all2=df_all2.reset_index(drop=True)

					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					df_all3=df_all.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
					df_all3=df_all3.reset_index(drop=True)

					df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))


					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
					totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

					total_hour=(startb-starthr)

					
					df_a=df_all1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_a1=df_a.resample(freqhr).nunique()
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_b1=df_b.resample(freqhr).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					df_c=df_all3[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_c1=df_c.resample(freqhr).nunique()
					df_c1.loc[df_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_c=df_c1["uuid_id"].sum()

					tot_org_c=(7*total_hour*totqr_c)/freq

					succrate_c=(tot_succ_c/tot_org_c)*100

					ovr_succ=(succrate_a+succrate_b+succrate_c)/3

					success_ls_n.append(round(succrate_c,2))
					success_ls.append(round(ovr_succ,2))



			else:
				df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_all["hour"]=df_all["time"].dt.hour

				df_old['time'] = pd.to_datetime(df_old['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')
				df_old["hour"]=df_old["time"].dt.hour

				freqhr=str(freq)+"H"
				offseta=str(starthr)+"H"
				offsetb=str(startb)+"H"
				offsetc=str(startc)+"H"


				if startc==0:
					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)

					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)|(df_all["hour"]<starthr)]
					df_all2=df_all2.reset_index(drop=True)

					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])

					total_hour=(startb-starthr)
					print("Total Hour",total_hour)

					
					df_a=df_all1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					
					df_a1=df_a.resample(freqhr).nunique()
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_b1=df_b.resample(freqhr,offset=offsetb).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					# if ele=="Emami-Amingaon":
					# 	print(df_all2["time"])
					# 	print(df_b1.tail(6))
					# 	print("This is sampling freq",freqhr)
					# 	print(startb)



					

					tot_org_b=(7*total_hour*totqr_b)/freq

					succrate_b=(tot_succ_b/tot_org_b)*100

					if ele=="Supreme Industries" or ele=="Supreme Industries SPRM01" or ele=="YORK PRINT YPRT03" or ele=="YORK PRINT YPRT04":
						ovr_succ=succrate_a
						success_ls.append(round(ovr_succ,2))
						success_ls_n.append(0)
					else:
						ovr_succ=(succrate_a+succrate_b)/2
						success_ls_n.append(round(succrate_b,2))
						success_ls.append(round(ovr_succ,2))

					# ovr_succ=(succrate_a+succrate_b)/2

					# success_ls_n.append(round(succrate_b,2))
					# success_ls.append(round(ovr_succ,2))


					df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
					df_old1=df_old1.reset_index(drop=True)
					df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

					df_old2=df_old.loc[(df_old["hour"]>=startb)|(df_old["hour"]<starthr)]
					df_old2=df_old2.reset_index(drop=True)
					df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

					total_hour=(startb-starthr)

					
					df_old_a=df_old1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_old_a1=df_old_a.resample(freqhr).nunique()
					# print(type(df_old_a1))
					df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_old_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_old_b=df_old2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_old_b1=df_old_b.resample(freqhr).nunique()
					df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_old_b1["uuid_id"].sum()

					tot_org_b=(7*total_hour*totqr_b)/freq

					succrate_b=(tot_succ_b/tot_org_b)*100

					if ele=="Supreme Industries" or ele=="Supreme Industries SPRM01" or ele=="YORK PRINT YPRT03" or ele=="YORK PRINT YPRT04":
						ovr_succ=succrate_a
						suc_old.append(round(ovr_succ,2))
						suc_old_n.append(0)
					else:
						# ovr_succ=(succrate_a+succrate_b)/2
						# success_ls_n.append(round(succrate_b,2))
						# success_ls.append(round(ovr_succ,2))

						ovr_succ=(succrate_a+succrate_b)/2

						suc_old_n.append(round(succrate_b,2))
						suc_old.append(round(ovr_succ,2))

				else:

					df_all1=df_all.loc[(df_all["hour"]>=starthr)&(df_all["hour"]<startb)]
					df_all1=df_all1.reset_index(drop=True)

					df_all1=df_all1.set_index(pd.to_datetime(df_all1['time']))

					df_all2=df_all.loc[(df_all["hour"]>=startb)&(df_all["hour"]<startc)]
					df_all2=df_all2.reset_index(drop=True)

					df_all2=df_all2.set_index(pd.to_datetime(df_all2['time']))

					df_all3=df_all3.loc[(df_all["hour"]>=startc)|(df_all["hour"]<starthr)]
					df_all3=df_all3.reset_index(drop=True)

					df_all3=df_all3.set_index(pd.to_datetime(df_all3['time']))

					totqr_a=int(df1.loc[df1["siteid"]==ele,"AQ"].tolist()[0])
					totqr_b=int(df1.loc[df1["siteid"]==ele,"BQ"].tolist()[0])
					totqr_c=int(df1.loc[df1["siteid"]==ele,"CQ"].tolist()[0])

					total_hour=(startb-starthr)

					
					df_a=df_all1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_a1=df_a.resample(freqhr).nunique()
					df_a1.loc[df_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_b=df_all2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_b1=df_b.resample(freqhr,offset=offsetb).nunique()
					df_b1.loc[df_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_b1["uuid_id"].sum()

					tot_org_b=(7*total_hour*totqr_b)/freq

					succrate_b=(tot_succ_b/tot_org_b)*100

					df_c=df_allc[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_c1=df_c.resample(freqhr,offset=offsetc).nunique()
					df_c1.loc[df_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_c=df_c1["uuid_id"].sum()

					tot_org_c=(7*total_hour*totqr_c)/freq

					succrate_c=(tot_succ_c/tot_org_c)*100

					ovr_succ=(succrate_a+succrate_b+succrate_c)/3

					success_ls_n.append(round(succrate_c,2))
					success_ls.append(round(ovr_succ,2))


					df_old1=df_old.loc[(df_old["hour"]>=starthr)&(df_old["hour"]<startb)]
					df_old1=df_old1.reset_index(drop=True)
					df_old1=df_old1.set_index(pd.to_datetime(df_old1['time']))

					df_old2=df_old.loc[(df_old["hour"]>=startb)&(df_old["hour"]<startc)]
					df_old2=df_old2.reset_index(drop=True)
					df_old2=df_old2.set_index(pd.to_datetime(df_old2['time']))

					df_old3=df_old.loc[(df_old["hour"]>=startc)|(df_old["hour"]<starthr)]
					df_old3=df_old3.reset_index(drop=True)
					df_old3=df_old3.set_index(pd.to_datetime(df_old3['time']))

					total_hour=(startb-starthr)

					
					df_old_a=df_old1[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_old_a1=df_old_a.resample(freqhr).nunique()
					# print(type(df_old_a1))
					df_old_a1.loc[df_old_a1["uuid_id"]>totqr_a,"uuid_id"]=totqr_a
					tot_succ_a=df_old_a1["uuid_id"].sum()

					tot_org_a=(7*total_hour*totqr_a)/freq

					succrate_a=(tot_succ_a/tot_org_a)*100

					df_old_b=df_old2[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_old_b1=df_old_b.resample(freqhr).nunique()
					df_old_b1.loc[df_old_b1["uuid_id"]>totqr_b,"uuid_id"]=totqr_b
					tot_succ_b=df_old_b1["uuid_id"].sum()

					tot_org_b=(7*total_hour*totqr_b)/freq

					succrate_b=(tot_succ_b/tot_org_b)*100

					df_old_c=df_oldc[["time","uuid_id"]]
					freqhr=str(freq)+"H"
					df_old_c1=df_old_c.resample(freqhr).nunique()
					df_old_c1.loc[df_old_c1["uuid_id"]>totqr_c,"uuid_id"]=totqr_c
					tot_succ_c=df_old_c1["uuid_id"].sum()

					tot_org_c=(7*total_hour*totqr_c)/freq

					succrate_c=(tot_succ_c/tot_org_c)*100

					ovr_succ=(succrate_a+succrate_b+succrate_c)/3

					suc_old_n.append(round(succrate_c,2))
					suc_old.append(round(ovr_succ,2))













			# print(start_date1,end_date1)
	df1["success"]=success_ls
	df1["successn"]=success_ls_n
	df1["oldsucc"]=suc_old
	df1["oldnight"]=suc_old_n
	df1["overdelta"]=round(df1["success"]-df1["oldsucc"],2)
	df1["nightd"]=round(df1["successn"]-df1["oldnight"],2)

	

	df1["color1"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["overdelta"]]
	df1["color2"]=["color:#00ff00" if x>=0 else "color:#ff0000" for x in df1["nightd"]]

	df1["arrow1"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["overdelta"]]
	df1["arrow2"]=["fa fa-arrow-up" if x>=0 else "fa fa-arrow-down" for x in df1["nightd"]]




	






	ovr_success=round(df1["success"].mean(),2)
	ovr_old=round(df1["oldsucc"].mean(),2)

	delta_ovr=round(ovr_success-ovr_old,2)

	if delta_ovr<0:
		cl1="fa fa-arrow-down"
		col1="color:#ff0000"
	else:
		cl1="fa fa-arrow-up"
		col1="color:#00ff00"


	if branch=="Assam":
		sum_night1=df1["successn"].sum()
		ovr_night=round(sum_night1/(len(df1)-4),2)
		sum_night_old=df1["oldnight"].sum()
		ovr_night_old=round(sum_night_old/(len(df1)-4),2)


	else:


		ovr_night=round(df1["successn"].mean(),2)
		ovr_night_old=round(df1["oldnight"].mean(),2)

	delta_night=round(ovr_night-ovr_night_old,2)

	if delta_night<0:
		cl2="fa fa-arrow-down"
		col2="color:#ff0000"
	else:
		cl2="fa fa-arrow-up"
		col2="color:#00ff00"

	# print(df1)
	# print(df1[["success","successn","oldsucc","oldnight","overdelta"]])













	json_records = df1.reset_index().to_json(orient ='records')
	data = []
	data = json.loads(json_records)

	# context={"branch":branch,"data":data}

	context={"branch":branch,"data":data,"ovr_s":ovr_success,"ovr_old":ovr_old,"ovr_n":ovr_night,"ovr_no":ovr_night_old,"del_ovr":round(delta_ovr,2),"del_n":round(delta_night,2),
	"cl1":cl1,"col1":col1,"cl2":cl2,"col2":col2,"dt1":dt1,"dt2":dt2,"totsite":len(df1)}


	# shift_timeA=[]
	# shift_timeB=[]
	# shift_timeC=[]
	# df1["A_freq"]=df1["Shift A freq"].astype
	# df1["shift_timeA"]=df1["Shift A"]+"\n"+df1["Shift A freq"].astype(str)+" minutes"
	# print(df1)

	# return HttpResponse("Success")

	return render(request,"patrolist.html",context)

def sitedata(request):
	site_code=request.GET.get("sitecode","")
	a=request.GET.get("A","")
	b=request.GET.get("B","")
	c=request.GET.get("C","")
	freq=request.GET.get("freq","")
	night=request.GET.get("night","")
	# af=request.GET.get("AF","")
	# bf=request.GET.get("BF","")
	# cf=request.GET.get("CF","")
	date_inp=request.GET.get("date_inp","")
	placeselect=request.GET.get("placeselect","")
	authkey=request.GET.get("auth","")
	auth_org="TjWnZr4u"

	if site_code!="AMRI-Mukundpur" and authkey==auth_org:
		return HttpResponse("<h3>You are not authorized to see this dashboard</h3>")
	elif site_code=="AMRI-Mukundpur" and authkey!=auth_org:
		return HttpResponse("<h3>Unauthorized access denied!</h3>")
	else:
		pass





	# print(af,bf,cf)

	if site_code=="AMRI-Mukundpur":
		a=7
		b=14
		c=21
		freq=30
		night=0



		if float(a)<10:
			a1="0"+str(a)
		else:
			a1=str(a)
		if float(b)<10:
			b1="0"+str(b)
		else:
			b1=str(b)
		if float(c)<10 and float(c)!=0:
			c1="0"+str(c)
		else:
			c1=str(c)

		siteinfo=[]

		if float(night)==1:
			sitetyp="onlynight"
		else:
			sitetyp="all"


	else:
		if float(a)<10:
			a1="0"+a
		else:
			a1=a
		if float(b)<10:
			b1="0"+b
		else:
			b1=b
		if float(c)<10 and float(c)!=0:
			c1="0"+c
		else:
			c1=c

		siteinfo=[]

		if float(night)==1:
			sitetyp="onlynight"
		else:
			sitetyp="all"





	# if (af=="Not observerd") or (bf=="Not observerd"):
	# 	print("program at first")
	# 	sitetyp="only night"
	# 	if (af=="Not observerd") and (bf!="Not observerd"):
	# 		if float(bf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	elif (af=="Not observerd" or "Not Applicable") or (bf=="Not observerd" or "Not Applicable"):
	# 		if float(cf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	else:
	# 		gap="30mnt"


	# else:
	# 	print("program is here")

	# 	if float(af)==60 and float(bf)==30:
	# 		sitetyp="all"
	# 		gap="mixed"

	# 	else:
	# 		sitetyp="all"
	# 		gap="30mnt"
	# siteinfo.append("Site Name: "+site_code)

	tz = timezone.get_current_timezone()

	if date_inp=="":


		time1=datetime.now()
	else:
		time1=datetime.strptime(date_inp, '%Y-%m-%d')
		time1=time1+timedelta(days=1,hours=a)



	if float(freq)==60:
		gap="1hr"
	elif float(freq)==30:
		gap="30min"
	elif float(freq)==45:
		gap="mixed"
	else:
		gap=="30min"

	if float(c)==0:
		shift_count=2
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+a1+":00 Hours")

		if gap=="1hr":
			tot_slot=24*1
		elif gap=="30min":
			tot_slot=24*2
		elif gap=="mixed":
			sftatime=float(b)-float(a)
			sftbtime=(24-float(b))+float(a)
			tot_slot=(sftatime*1)+(sftbtime*2)
			print("this is total slot",tot_slot)
		else:
			tot_slot=24*2

		tot_slot=24






	else:
		shift_count=3
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+c1+":00 Hours")
		siteinfo.append("SHIFT C:"+c1+":00-"+a1+":00 Hours")
		tot_slot=24


	if sitetyp=="onlynight":
		siteinfo.append("Monitoring happens only in Night Shift")
	else:
		siteinfo.append("Monitoring happens 24 Hours")

	if float(freq)==60:
		gap="1hr"
		siteinfo.append("Post Monitoring Frequency: Every 1 Hour")
	elif gap=="mixed":
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes in Shift B. Every 1 hour in Shift A.")
	else:
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes")
	


	
	# time1=tz.localize(time1
	# time1=time1+timedelta(hours=5,minutes=30)
	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time2_1+timedelta(hours=int(float(a)))
	time2_11=time2_1-timedelta(days=1)


	# endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time2_1)
	start_time1=tz.localize(time2_11)

	starthr=int(start_time1.hour)
	print(starthr)


	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code,uuidtype="QR").values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

	siteinfo.append("Total No. of Posts: "+str(len(all_beacons_names)))


	dict_hour={"time":[],"hr":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"\n\n"+all_beacons_names["placeid_id"]

	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	for ele in name_list:
		dict_hour[ele]=[]

	if len(all_path)==0:
		data_avl=False
		start_hr=starthr
		if starthr==23:
			end_hr=0
		else:
			end_hr=starthr+1
		i=0
		now_hour=int(time1.hour)

		# total_slot=13

		total_slot=tot_slot



		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	otal_slot=abs(now_hour-start_hr+1)

		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:

			print(starthr)



			# if start_hr==23:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
			# else:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
			# print(len(df_all_1))


			# if start_hr<10:





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)



			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			for idx,ele1 in enumerate(name_list):
				# counter=counter+1
				
				# print("program is here")
				if gap=="1hr":
					dict_hour[ele1].append("color:#ff0000")

				elif gap=="mixed":
					if start_hr>=float(a) and start_hr<float(b):
						dict_hour[ele1].append("color:#ff0000")
					else:
						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")

				else:

					dict_hour[ele1].append("color:#ff0000")
					dict_hour[ele1].append("color:#ff0000")

			if start_hr==23:
				start_hr=0
			else:
				start_hr=start_hr+1

			if end_hr==23:
				end_hr=0

			else:
				end_hr=end_hr+1

			i=i+1


	else:
		data_avl=True
		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute

		df_all_tbl=df_all[["fullname","time","datetime1","hour","empname"]]

		# print(df_all["datetime1"])


		start_hr=starthr
		if starthr==23:
			end_hr=0
		else:
			end_hr=starthr+1
		i=0
		now_hour=int(time1.hour)

		total_slot=tot_slot



		# start_hr=starthr
		# end_hr=starthr+1


		# i=0
		# now_hour=int(time1.hour)

		# total_slot=abs(now_hour-start_hr+1)
		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	total_slot=abs(now_hour-start_hr+1)
		# total_slot=abs(now_hour-start_hr+1)
		# print("this is total slot", total_slot)
		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:


			if start_hr==23:
				df_all_1=df_all.loc[(df_all["hour"]>=start_hr)| ((df_all["hour"]==0)&(df_all["minute"]==0))]
			else:
				# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
				df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]
			# print(len(df_all_1))


			# if start_hr<10:

			# print(starthr)





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)
			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			if len(df_all_1)==0:
				
				for idx,ele1 in enumerate(name_list):
					# counter=counter+1
					
					# print("program is here")
					if gap=="1hr":
						dict_hour[ele1].append("color:#ff0000")

					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							dict_hour[ele1].append("color:#ff0000")
						else:
							dict_hour[ele1].append("color:#ff0000")
							dict_hour[ele1].append("color:#ff0000")
					else:

						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")



			else:
				


				for idx1,ele2 in enumerate(name_list):
					# counter=counter+1

					if gap=="1hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:

							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")

					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]
							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:

								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
						else:

							df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
								# counter=counter+1


							df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_3)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_3["datetime1"].tolist()[-1]

								empn=df_all_3["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")





					else:

					
					
						df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
							# counter=counter+1


						df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

						if len(df_all_3)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_3["datetime1"].tolist()[-1]

							empn=df_all_3["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
						# counter=counter+1


			if start_hr==23:
				start_hr=0
			else:
				start_hr=start_hr+1

			if end_hr==23:
				end_hr=0

			else:
				end_hr=end_hr+1

			i=i+1


	df_final=pd.DataFrame(dict_hour)

	df_final.drop('hr', axis=1, inplace=True)

	# hrls=dict_hour["hr"]
	# sftls=[]

	# for ele3 in hrls:
	# 	if float(c)==0:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		else:
	# 			sftls.append("B")

	# 	else:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		elif ele3>=float(b) and ele3<float(c):
	# 			sftls.append("B")
	# 		else:
	# 			sftls.append("C")


	# df_final["sft"]=sftls




	# df_final["sft"]=["A" if (x>=float(a)and x<float(b))   else "color:#ff0000" for x in df_final["hr"]]



	# print(counter)
	# print(len(dict_hour))

	# for key in dict_hour:
	# 	print(len(dict_hour[key]),key)

	name_list1=[]




	for ele10 in name_list:
		x1=ele10[0:(len(ele10)-len(site_code)-2)]
		name_list1.append(x1)


	if data_avl==True:

		hr_list=df_all_tbl["hour"].tolist()
		sh_list=[]

		for ele12 in hr_list:
			if ele12>=a and ele12<b:
				sh_list.append("A")
			elif ele12>=b and ele12<c:
				sh_list.append("B")
			else:
				sh_list.append("C")

		df_all_tbl["shift"]=sh_list
		# df_all_tbl["name_ls"]=name_list1

		df_all_tbl=df_all_tbl.sort_values("time",ascending=True)
		json_records4 = df_all_tbl.reset_index().to_json(orient ='records')
		data4 = []
		data4 = json.loads(json_records4)
	else:
		data4=False


	# df_all_tbl["shift"]=["A" if x>=a and x<b elif(x>=b and x<c)  else "C" for x in df_all_tbl["hour"]]

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)

	


	records=df_final.to_records(index=False)
	result=tuple(records)
	# result=result[::-1]
	# print(result)







	context={"namelist":name_list1,"dt_1":dt_1,"data":result,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap, "siteinfo":siteinfo,
	"a":a,"b":b,"c":c,"freq":freq,"night":night,"filter_date":date_inp,"d":data4,"auth":authkey}


	return render(request,"sitedata.html",context)



def deviation(request):
	site_code=request.GET.get("sitecode","")
	a=request.GET.get("A","")
	b=request.GET.get("B","")
	c=request.GET.get("C","")
	freq=request.GET.get("freq","")
	night=request.GET.get("night","")
	placeselect=request.GET.get("placeselect","")
	date_inp=request.GET.get("date_inp","")

	tz = timezone.get_current_timezone()
	time1=datetime.now()


	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time2_1+timedelta(hours=int(float(a)))
	time2_11=time2_1-timedelta(days=1)


	# endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time2_1)
	start_time1=tz.localize(time2_11)

	starthr=int(start_time1.hour)

	gap=float(freq)
	# print(starthr)


	print(start_time1,end_time1)

	start_time1=start_time1-timedelta(days=1)
	end_time1=end_time1-timedelta(days=1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	print(dt_1)

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code,uuidtype="QR").values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()



	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

	# siteinfo.append("Total No. of Posts: "+str(len(all_beacons_names)))


	dict_hour={"time":[],"hr":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]

	if len(all_path)!=0:

		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute

		df_emp_grp=df_all.groupby("empname")

		data_list=[]
		time_list=[]
		datetime_list=[]
		# emp_list=[]

		for n1, grp1 in df_emp_grp:

			df_time_grp=grp1.groupby("datetime1")

			for n2,grp2 in df_time_grp:
				if len(grp2)>1:
					val_lis=grp2["fullname"].unique().tolist()


					name_list1=[]

					for ele10 in val_lis:
						x1=ele10[0:(len(ele10)-len(site_code)-2)]
						name_list1.append(x1)

					if len(val_lis)>1:
						time_list.append(grp2["time"].tolist()[0])
						str_val=" "
						str_val=", ".join(name_list1)
						data_list.append(n1+" scans "+str(len(val_lis))+ " places at the same time at "+ n2+": ")
						datetime_list.append(str_val)


		# print(data_list,time_list)


		df_emp_time=df_all.groupby("time")

		time_list1=[]
		data_list1=[]

		for n3,grp3 in df_emp_time:
			if len(grp3)>1:
				val_lis=grp2["fullname"].unique().tolist()
				if len(val_lis)>4:
					name_list1=[]

					for ele10 in val_lis:
						x1=ele10[0:(len(ele10)-len(site_code)-2)]
						name_list1.append(x1)

					str_val=" "
					str_val=",".join(name_list1)
					data_list1.append(str(len(val_lis))+" places are scanned at the same time, at "+grp3["datetime1"].tolist()[0])
					time_list1.append(n3)


		# print(data_list1)



	start_time11=start_time1-timedelta(days=5)
	end_time11=end_time1-timedelta(days=5)

	all_path1=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time11,end_time11)).values()))

	df_uuid_grp=all_path1.groupby("uuid_id",as_index = False).agg(mean_lat=pd.NamedAgg(column="lat",aggfunc="mean"),mean_long=pd.NamedAgg(column="long",aggfunc="mean"),
		std_long=pd.NamedAgg(column="long",aggfunc="std"),std_lat=pd.NamedAgg(column="lat",aggfunc="std"))
	print(df_uuid_grp)

	data_list2=[]
	time_list2=[]

	for index,rows in df_uuid_grp.iterrows():
		df_all_new=df_all.loc[df_all["uuid_id"]==rows["uuid_id"]]

		if len(df_all_new)>0:
			df_all_new1=df_all_new.reset_index()

			for ind1,row1 in df_all_new1.iterrows():

				if abs(row1["lat"]-rows["mean_lat"])>3*rows["std_lat"] or abs(row1["long"]-rows["mean_long"])>3*rows["std_long"]:
					time_list2.append(row1["time"])
					data_line="deviation occured at "+row1["datetime1"]+" by "+row1["empname"]+" at "+row1["name"]
					data_list2.append(data_line)
					# print(abs(row1["lat"]-rows["mean_lat"]),rows["std_lat"])




	# print(len(data_list2))



	if len(data_list)>0:
		df_1=pd.DataFrame({"time":time_list,"data":data_list,"place":datetime_list})
		df_1=df_1.sort_values("time",ascending=True)

		json_records1 = df_1.reset_index().to_json(orient ='records')
		data1 = []
		data1 = json.loads(json_records1)

	else:
		data1=False


	if len(data_list1)>0:
		df_2=pd.DataFrame({"time":time_list1,"data":data_list1})
		df_2=df_2.sort_values("time",ascending=True)

		json_records2 = df_2.reset_index().to_json(orient ='records')
		data2 = []
		data2 = json.loads(json_records2)
	else:
		data2=False

	if len(data_list2)>0:
		df_3=pd.DataFrame({"time":time_list2,"data":data_list2})
		df_3=df_3.sort_values("time",ascending=True)

		json_records3 = df_3.reset_index().to_json(orient ='records')
		data3 = []
		data3 = json.loads(json_records3)
	else:
		data3=False


	context={"dt_1":dt_1,"d1":data1,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap,
	"a":a,"b":b,"c":c,"freq":freq,"night":night,"filter_date":date_inp,"d3":data3,"d2":data2}


	return render(request,"deviation.html",context)


















def livepostorg(request):
	site_code=request.GET.get("sitecode","")
	a=request.GET.get("A","")
	b=request.GET.get("B","")
	c=request.GET.get("C","")
	freq=request.GET.get("freq","")
	night=request.GET.get("night","")
	dayonly=request.GET.get("day","")
	print(dayonly)
	placeselect=request.GET.get("placeselect","")

	if float(a)<10:
		a1="0"+a
	else:
		a1=a
	if float(b)<10:
		b1="0"+b
	else:
		b1=b
	if float(c)<10 and float(c)!=0:
		c1="0"+c
	else:
		c1=c

	siteinfo=[]

	if float(night)==1:
		sitetyp="onlynight"
	elif dayonly!="" and float(dayonly)==1:
		sitetyp="onlyday"
		print("program is here")
	else:
		sitetyp="all"





	tz = timezone.get_current_timezone()
	time1=datetime.now()

	# if date_inp=="":


	# 	time1=datetime.now()
	# else:
	# 	time1=datetime.strptime(date_inp, '%Y-%m-%d')
	# 	time1=time1+timedelta(days=1,hours=6)



	if float(freq)==60:
		gap="1hr"
	elif float(freq)==30:
		gap="30min"
	elif float(freq)==45:
		gap="mixed"
	elif float(freq)==120:
		gap="2hr"
	else:
		gap=="30min"

	if float(c)==0:
		shift_count=2
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+a1+":00 Hours")

		if gap=="1hr":
			tot_slot=24*1
		elif gap=="30min":
			tot_slot=24*2
		elif gap=="mixed":
			sftatime=float(b)-float(a)
			sftbtime=(24-float(b))+float(a)
			tot_slot=(sftatime*1)+(sftbtime*2)
			print("this is total slot",tot_slot)
		else:
			tot_slot=24*2

		tot_slot=12






	else:
		shift_count=3
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+c1+":00 Hours")
		siteinfo.append("SHIFT C:"+c1+":00-"+a1+":00 Hours")
		tot_slot=12


	if sitetyp=="onlynight":
		siteinfo.append("Monitoring happens only in Night Shift")
	elif sitetyp=="onlyday":
		siteinfo.append("Monitoring happens only between "+":"+a1+":00 to"+b1+":00 Hours")
	else:
		siteinfo.append("Monitoring happens 24 Hours")

	if float(freq)==60:
		gap="1hr"
		siteinfo.append("Post Monitoring Frequency: Every 1 Hour")
	elif gap=="mixed":
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes in Shift B. Every 1 hour in Shift A.")
	elif gap=="2hr":
		siteinfo.append("Post Monitoring Frequency: Every 2 Hour")
	else:
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes")
	


	
	# time1=tz.localize(time1
	# time1=time1+timedelta(hours=5,minutes=30)
	# time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time1
	# time2_1=time2_1+timedelta(hours=int(float(a)))
	time2_11=time2_1-timedelta(hours=12)


	# endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time2_1)
	start_time1=tz.localize(time2_11)

	starthr=int(start_time1.hour)
	# print(starthr)


	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code,uuidtype="QR").values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

	siteinfo.append("Total No. of Posts: "+str(len(all_beacons_names)))


	dict_hour={"time":[],"hr":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"-"+all_beacons_names["placeid_id"]
	# print(all_beacons_names["ful"])

	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	end_time2=end_time1
	start_time2=end_time2-timedelta(hours=2)

	dt_2=datetime.strftime(start_time2,"%d-%B %I:%M %p") +" to "+ datetime.strftime(end_time2,"%d-%B %I:%M %p")

	# print("This is time:", start_time2,end_time2)



	df_last2=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time2,end_time2)).values()))

	if len(df_last2)==0:
		name_last2_ls=name_list
	else:

		df_last1=pd.merge(df_last2,all_beacons_names,on='uuid_id')
		name_avl=df_last1["fullname"].tolist()

		# print(name_avl)

		name_last2_ls=[]
		# avail_shift=[]

		for ele11 in name_list:
			if ele11 not in name_avl:
				name_last2_ls.append(ele11)


	for ele in name_list:
		dict_hour[ele]=[]

	if len(all_path)==0:
		start_hr=starthr

		if gap!="2hr":
			if starthr==23:
				end_hr=0
			else:
				end_hr=starthr+1
			i=0
			now_hour=int(time1.hour)

			# total_slot=13


			total_slot=tot_slot
		else:
			if starthr==22:
				end_hr=0
			else:
				end_hr=starthr+2
			i=0
			now_hour=int(time1.hour)

			# total_slot=13


			total_slot=tot_slot/2



		# print(total_slot)



		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	otal_slot=abs(now_hour-start_hr+1)

		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:

			print(starthr)



			# if start_hr==23:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
			# else:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
			# print(len(df_all_1))


			# if start_hr<10:





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="2hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)



			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			for idx,ele1 in enumerate(name_list):
				# counter=counter+1
				
				# print("program is here")
				if gap=="1hr":
					dict_hour[ele1].append("color:#ff0000")

				elif gap=="2hr":
					dict_hour[ele1].append("color:#ff0000")


				elif gap=="mixed":
					if start_hr>=float(a) and start_hr<float(b):
						dict_hour[ele1].append("color:#ff0000")
					else:
						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")

				else:

					dict_hour[ele1].append("color:#ff0000")
					dict_hour[ele1].append("color:#ff0000")

			if gap!="2hr":

				if start_hr==23:
					start_hr=0
				else:
					start_hr=start_hr+1

				if end_hr==23:
					end_hr=0

				else:
					end_hr=end_hr+1

				i=i+1
			else:
				if start_hr==22:
					start_hr=0
				else:
					start_hr=start_hr+2

				if end_hr==22:
					end_hr=0

				else:
					end_hr=end_hr+2

				i=i+1


	else:
		# print("program is here")
		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute

		# print(df_all["datetime1"])


		start_hr=starthr

		if gap!="2hr":
			if starthr==23:
				end_hr=0
			else:
				end_hr=starthr+1
			i=0
			now_hour=int(time1.hour)

			# total_slot=13


			total_slot=tot_slot
		else:
			if starthr==22:
				end_hr=0
			else:
				end_hr=starthr+2
			i=0
			now_hour=int(time1.hour)

			# total_slot=13


			total_slot=tot_slot/2
		# if starthr==23:
		# 	end_hr=0
		# else:
		# 	end_hr=starthr+1
		# i=0
		# now_hour=int(time1.hour)

		# total_slot=tot_slot

		# print(start)



		# start_hr=starthr
		# end_hr=starthr+1


		# i=0
		# now_hour=int(time1.hour)

		# total_slot=abs(now_hour-start_hr+1)
		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	total_slot=abs(now_hour-start_hr+1)
		# total_slot=abs(now_hour-start_hr+1)
		# print("this is total slot", total_slot)
		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:
			# print("Program in while loop")

			# print(start_hr,end_hr)


			if gap!="2hr":


				if start_hr==23:
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr)| ((df_all["hour"]==0)&(df_all["minute"]==0))]
				else:
					# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]

			else:
				if start_hr==22:
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr)| ((df_all["hour"]==0)&(df_all["minute"]==0))]
				else:
					# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]

			# print(len(df_all_1))


			# if start_hr<10:

			# print(starthr)





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="2hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)


			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)
			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			if len(df_all_1)==0:
				
				for idx,ele1 in enumerate(name_list):
					# counter=counter+1
					# print("program is in for loop 1")

					
					# print("program is here")
					if gap=="1hr":
						dict_hour[ele1].append("color:#ff0000")

					elif gap=="2hr":
						dict_hour[ele1].append("color:#ff0000")

					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							dict_hour[ele1].append("color:#ff0000")
						else:
							dict_hour[ele1].append("color:#ff0000")
							dict_hour[ele1].append("color:#ff0000")
					else:

						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")



			else:
				


				for idx1,ele2 in enumerate(name_list):
					# counter=counter+1
					# print("program is in for loop 2")

					if gap=="1hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:

							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")

					elif gap=="2hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:

							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")


					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]
							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:

								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
						else:

							df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
								# counter=counter+1


							df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_3)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_3["datetime1"].tolist()[-1]

								empn=df_all_3["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")





					else:

					
					
						df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
							# counter=counter+1


						df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

						if len(df_all_3)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_3["datetime1"].tolist()[-1]

							empn=df_all_3["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
						# counter=counter+1


			if gap!="2hr":

				if start_hr==23:
					start_hr=0
				else:
					start_hr=start_hr+1

				if end_hr==23:
					end_hr=0

				else:
					end_hr=end_hr+1

				i=i+1
			else:
				if start_hr==22:
					start_hr=0
				else:
					start_hr=start_hr+2

				if end_hr==22:
					end_hr=0

				else:
					end_hr=end_hr+2

				i=i+1


	df_final=pd.DataFrame(dict_hour)

	df_final.drop('hr', axis=1, inplace=True)

	# hrls=dict_hour["hr"]
	# sftls=[]

	# for ele3 in hrls:
	# 	if float(c)==0:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		else:
	# 			sftls.append("B")

	# 	else:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		elif ele3>=float(b) and ele3<float(c):
	# 			sftls.append("B")
	# 		else:
	# 			sftls.append("C")


	# df_final["sft"]=sftls




	# df_final["sft"]=["A" if (x>=float(a)and x<float(b))   else "color:#ff0000" for x in df_final["hr"]]



	# print(counter)
	# print(len(dict_hour))

	# for key in dict_hour:
	# 	print(len(dict_hour[key]),key)

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)


	records=df_final.to_records(index=False)
	result=tuple(records)
	result=result[::-1]
	# print(result)

	name_list1=[]




	for ele10 in name_list:
		x1=ele10[0:(len(ele10)-len(site_code)-2)]
		name_list1.append(x1)

	print(name_last2_ls,dt_2)





	context={"namelist":name_list1,"dt_1":dt_1,"data":result,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap, "siteinfo":siteinfo,
	"a":a,"b":b,"c":c,"freq":freq,"night":night,"name_last2_ls":name_last2_ls}


	return render(request,"livepost.html",context)



def live(request):

	site_code=request.GET.get("sitecode","")
	a=request.GET.get("A","")
	b=request.GET.get("B","")
	c=request.GET.get("C","")
	freq=request.GET.get("freq","")
	night=request.GET.get("night","")
	dayonly=request.GET.get("day","")
	# af=request.GET.get("AF","")
	# bf=request.GET.get("BF","")
	# cf=request.GET.get("CF","")
	date_inp=request.GET.get("date_inp","")
	placeselect=request.GET.get("placeselect","")


	# print(af,bf,cf)



	if float(a)<10:
		a1="0"+a
	else:
		a1=a
	if float(b)<10:
		b1="0"+b
	else:
		b1=b
	if float(c)<10 and float(c)!=0:
		c1="0"+c
	else:
		c1=c

	siteinfo=[]

	if float(night)==1:
		sitetyp="onlynight"
	elif float(dayonly)==1:
		sitetyp="onlyday"
	else:
		sitetyp="all"




	# if (af=="Not observerd") or (bf=="Not observerd"):
	# 	print("program at first")
	# 	sitetyp="only night"
	# 	if (af=="Not observerd") and (bf!="Not observerd"):
	# 		if float(bf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	elif (af=="Not observerd" or "Not Applicable") or (bf=="Not observerd" or "Not Applicable"):
	# 		if float(cf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	else:
	# 		gap="30mnt"


	# else:
	# 	print("program is here")

	# 	if float(af)==60 and float(bf)==30:
	# 		sitetyp="all"
	# 		gap="mixed"

	# 	else:
	# 		sitetyp="all"
	# 		gap="30mnt"
	# siteinfo.append("Site Name: "+site_code)

	tz = timezone.get_current_timezone()

	if date_inp=="":


		time1=datetime.now()
	else:
		time1=datetime.strptime(date_inp, '%Y-%m-%d')
		time1=time1+timedelta(days=1,hours=6)



	if float(freq)==60:
		gap="1hr"
	elif float(freq)==120:
		gap="2hr"
	elif float(freq)==30:
		gap="30min"
	elif float(freq)==45:
		gap="mixed"
	else:
		gap=="30min"

	if float(c)==0 :
		shift_count=2
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+a1+":00 Hours")

		if gap=="1hr":
			tot_slot=24*1
		elif gap=="30min":
			tot_slot=24*2
		elif gap=="mixed":
			sftatime=float(b)-float(a)
			sftbtime=(24-float(b))+float(a)
			tot_slot=(sftatime*1)+(sftbtime*2)
			print("this is total slot",tot_slot)
		else:
			tot_slot=24*2

		tot_slot=24






	else:
		shift_count=3
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+c1+":00 Hours")
		siteinfo.append("SHIFT C:"+c1+":00-"+a1+":00 Hours")
		tot_slot=24


	if sitetyp=="onlynight":
		siteinfo.append("Monitoring happens only in Night Shift")
	elif sitetyp=="onlyday":
		siteinfo.append("Monitoring happens only between "+":"+a1+":00 to "+b1+":00 Hours")
	else:
		siteinfo.append("Monitoring happens 24 Hours")

	if float(freq)==60:
		gap="1hr"
		siteinfo.append("Post Monitoring Frequency: Every 1 Hour")
	elif float(freq)==120:
		siteinfo.append("Post Monitoring Frequency: Every 2 Hour")
	elif gap=="mixed":
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes in Shift B. Every 1 hour in Shift A.")
	else:
		siteinfo.append("Post Monitoring Frequency: Every 30 minutes")
	


	
	# time1=tz.localize(time1
	# time1=time1+timedelta(hours=5,minutes=30)
	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time2_1+timedelta(hours=int(float(a)))
	time2_11=time2_1-timedelta(days=1)


	# endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time2_1)
	start_time1=tz.localize(time2_11)

	starthr=int(start_time1.hour)
	print(starthr)


	print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code,uuidtype="QR").values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR",time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

	siteinfo.append("Total No. of Posts: "+str(len(all_beacons_names)))


	dict_hour={"time":[],"hr":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"\n\n"+all_beacons_names["placeid_id"]

	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	for ele in name_list:
		dict_hour[ele]=[]

	if len(all_path)==0:
		data_avl=False
		start_hr=starthr
		if starthr==23:
			end_hr=0
		else:
			end_hr=starthr+1
		i=0
		now_hour=int(time1.hour)

		# total_slot=13

		if gap=="2hr":

			total_slot=tot_slot/2
		else:
			total_slot=tot_slot





		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	otal_slot=abs(now_hour-start_hr+1)

		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:

			print(starthr)



			# if start_hr==23:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr)]
			# else:
			# 	df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
			# print(len(df_all_1))


			# if start_hr<10:





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="2hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)


			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)



			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			for idx,ele1 in enumerate(name_list):
				# counter=counter+1
				
				# print("program is here")
				if gap=="1hr":
					dict_hour[ele1].append("color:#ff0000")
				elif gap=="2hr":
					dict_hour[ele1].append("color:#ff0000")


				elif gap=="mixed":
					if start_hr>=float(a) and start_hr<float(b):
						dict_hour[ele1].append("color:#ff0000")
					else:
						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")

				else:

					dict_hour[ele1].append("color:#ff0000")
					dict_hour[ele1].append("color:#ff0000")


			if gap!="2hr":

				if start_hr==23:
					start_hr=0
				else:
					start_hr=start_hr+1

				if end_hr==23:
					end_hr=0

				else:
					end_hr=end_hr+1

				i=i+1
			else:
				if start_hr==22:
					start_hr=0
				elif start_hr==23:
					start_hr=1
				else:
					start_hr=start_hr+2

				if end_hr==22:
					end_hr=0

				elif end_hr==23:
					end_hr==1

				else:
					end_hr=end_hr+2

				i=i+1



	else:
		data_avl=True
		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute

		df_all_tbl=df_all[["fullname","time","datetime1","hour","empname"]]

		# print(df_all["datetime1"])



		if gap!="2hr":


			start_hr=starthr
			if starthr==23:
				end_hr=0
			else:
				end_hr=starthr+1
			i=0
			now_hour=int(time1.hour)
		else:
			start_hr=starthr
			if starthr==22:
				end_hr=0
			else:
				end_hr=starthr+2
			i=0
			now_hour=int(time1.hour)


		if gap=="2hr":

			total_slot=tot_slot/2
		else:
			total_slot=tot_slot

		# total_slot=tot_slot



		# start_hr=starthr
		# end_hr=starthr+1


		# i=0
		# now_hour=int(time1.hour)

		# total_slot=abs(now_hour-start_hr+1)
		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	total_slot=abs(now_hour-start_hr+1)
		# total_slot=abs(now_hour-start_hr+1)
		# print("this is total slot", total_slot)
		counter=0

		uuid_list=all_beacons_names["uuid_id"].tolist()

		while i<total_slot:


			if gap!="2hr":


				if start_hr==23:
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr)| ((df_all["hour"]==0)&(df_all["minute"]==0))]
				else:
					# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]
			else:
				if start_hr==22:
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr)| ((df_all["hour"]==0)&(df_all["minute"]==0))]
				else:
					# df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr)]
					df_all_1=df_all.loc[(df_all["hour"]>=start_hr) & (df_all["hour"]<end_hr) | ((df_all["hour"]==end_hr)&(df_all["minute"]==0))]

			# print(len(df_all_1))


			# if start_hr<10:

			# print(starthr)





			if start_hr<10:
				strt="0"+str(start_hr)
			else:
				strt=str(start_hr)

			if end_hr<10:
				ent="0"+str(end_hr)
			else:
				ent=str(end_hr)

			if gap=="1hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="2hr":
				time_stamp1=strt+":00-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)

			elif gap=="mixed":
				if start_hr>=float(a) and start_hr<float(b):
					time_stamp1=strt+":00-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
				else:
					time_stamp1=strt+":00-"+strt+":30 Hours"
					dict_hour["time"].append(time_stamp1)
					dict_hour["hr"].append(start_hr)
					time_stamp2=strt+":30-"+ent+":00 Hours"
					dict_hour["time"].append(time_stamp2)
					dict_hour["hr"].append(start_hr)
			else:

				time_stamp1=strt+":00-"+strt+":30 Hours"
				dict_hour["time"].append(time_stamp1)
				dict_hour["hr"].append(start_hr)
				time_stamp2=strt+":30-"+ent+":00 Hours"
				dict_hour["time"].append(time_stamp2)
				dict_hour["hr"].append(start_hr)

			if len(df_all_1)==0:
				
				for idx,ele1 in enumerate(name_list):
					# counter=counter+1
					
					# print("program is here")
					if gap=="1hr":
						dict_hour[ele1].append("color:#ff0000")

					elif gap=="2hr":
						dict_hour[ele1].append("color:#ff0000")


					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							dict_hour[ele1].append("color:#ff0000")
						else:
							dict_hour[ele1].append("color:#ff0000")
							dict_hour[ele1].append("color:#ff0000")
					else:

						dict_hour[ele1].append("color:#ff0000")
						dict_hour[ele1].append("color:#ff0000")



			else:
				


				for idx1,ele2 in enumerate(name_list):
					# counter=counter+1

					if gap=="1hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:

							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")

					elif gap=="2hr":
						df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]
						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:

							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")


					elif gap=="mixed":
						if start_hr>=float(a) and start_hr<float(b):
							df_all_2=df_all_1.loc[(df_all_1["fullname"]==ele2)]
							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:

								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
						else:

							df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_2)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_2["datetime1"].tolist()[-1]

								empn=df_all_2["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")
								# counter=counter+1


							df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

							if len(df_all_3)==0:
								dict_hour[ele2].append("color:#ff0000")
								# counter=counter+1
							else:
								tmlist=df_all_3["datetime1"].tolist()[-1]

								empn=df_all_3["empname"].tolist()[-1]

								# dict_hour[ele2].append("color:#00ff00;")
								dict_hour[ele2].append(tmlist+"("+empn+")")





					else:

					
					
						df_all_2=df_all_1.loc[(df_all_1["minute"]>0) & (df_all_1["minute"]<=30)&(df_all_1["fullname"]==ele2)]



						if len(df_all_2)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_2["datetime1"].tolist()[-1]

							empn=df_all_2["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
							# counter=counter+1


						df_all_3=df_all_1.loc[(df_all_1["minute"]>30)&(df_all_1["fullname"]==ele2)]

						if len(df_all_3)==0:
							dict_hour[ele2].append("color:#ff0000")
							# counter=counter+1
						else:
							tmlist=df_all_3["datetime1"].tolist()[-1]

							empn=df_all_3["empname"].tolist()[-1]

							# dict_hour[ele2].append("color:#00ff00;")
							dict_hour[ele2].append(tmlist+"("+empn+")")
						# counter=counter+1

			if gap!="2hr":
				if start_hr==23:
					start_hr=0
				else:
					start_hr=start_hr+1

				if end_hr==23:
					end_hr=0

				else:
					end_hr=end_hr+1

				i=i+1

			else:
				if start_hr==22:
					start_hr=0
				elif start_hr==23:
					start_hr=1
				else:
					start_hr=start_hr+2

				if end_hr==22:
					end_hr=0

				elif end_hr==23:
					end_hr==1

				else:
					end_hr=end_hr+2

				i=i+1



	df_final=pd.DataFrame(dict_hour)

	df_final.drop('hr', axis=1, inplace=True)

	# hrls=dict_hour["hr"]
	# sftls=[]

	# for ele3 in hrls:
	# 	if float(c)==0:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		else:
	# 			sftls.append("B")

	# 	else:
	# 		if ele3>=float(a) and ele3<float(b):
	# 			sftls.append("A")
	# 		elif ele3>=float(b) and ele3<float(c):
	# 			sftls.append("B")
	# 		else:
	# 			sftls.append("C")


	# df_final["sft"]=sftls




	# df_final["sft"]=["A" if (x>=float(a)and x<float(b))   else "color:#ff0000" for x in df_final["hr"]]



	# print(counter)
	# print(len(dict_hour))

	# for key in dict_hour:
	# 	print(len(dict_hour[key]),key)

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)


	records=df_final.to_records(index=False)
	result=tuple(records)
	# result=result[::-1]
	# print(result)

	name_list1=[]




	for ele10 in name_list:
		x1=ele10[0:(len(ele10)-len(site_code)-2)]
		name_list1.append(x1)


	if data_avl==True:

		# hr_list=df_all_tbl["hour"].tolist()
		# sh_list=[]

		# for ele12 in hr_list:
		# 	if ele12>=a and ele12<b:
		# 		sh_list.append("A")
		# 	elif ele12>=b and ele12<c:
		# 		sh_list.append("B")
		# 	else:
		# 		sh_list.append("C")

		# df_all_tbl["shift"]=sh_list
		# df_all_tbl["name_ls"]=name_list1

		df_all_tbl=df_all_tbl.sort_values("time",ascending=True)
		json_records4 = df_all_tbl.reset_index().to_json(orient ='records')
		data4 = []
		data4 = json.loads(json_records4)
	else:
		data4=False





	context={"namelist":name_list1,"dt_1":dt_1,"data":result,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap, "siteinfo":siteinfo,
	"a":a,"b":b,"c":c,"freq":freq,"night":night,"filter_date":date_inp,"d":data4,"data_avl":data_avl,"day":dayonly}


	return render(request,"livedata.html",context)


def livepatrol(request):

	site_code=request.GET.get("sitecode","")
	a=request.GET.get("A","")
	b=request.GET.get("B","")
	c=request.GET.get("C","")
	af=request.GET.get("AF","")
	bf=request.GET.get("BF","")
	cf=request.GET.get("CF","")
	placeselect=request.GET.get("placeselect","")
	print(af,bf,cf)

	if float(a)<10:
		a1="0"+str(int(float(a)))
	else:
		a1=str(int(float(a)))
	if float(b)<10:
		b1="0"+str(int(float(b)))
	else:
		b1=str(int(float(b)))
	if float(c)<10 and float(c)!=0:
		c1="0"+str(int(float(c)))
	else:
		c1=str(int(float(b)))

	siteinfo=[]


	gap=float(bf)

	tot_slot=24/gap

	# if (af=="Not observerd") or (bf=="Not observerd"):
	# 	print("program at first")
	# 	sitetyp="only night"
	# 	if (af=="Not observerd") and (bf!="Not observerd"):
	# 		if float(bf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	elif (af=="Not observerd" or "Not Applicable") or (bf=="Not observerd" or "Not Applicable"):
	# 		if float(cf)==60:
	# 			gap="1hr"
	# 		else:
	# 			gap="30mnt"
	# 	else:
	# 		gap="30mnt"


	# else:
	# 	print("program is here")

	# 	if float(af)==60 and float(bf)==30:
	# 		sitetyp="all"
	# 		gap="mixed"

	# 	else:
	# 		sitetyp="all"
	# 		gap="30mnt"
	siteinfo.append("Site Name: "+site_code)

	if float(c)==0:
		shift_count=2
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+a1+":00 Hours")


	else:
		shift_count=3
		siteinfo.append("No. of Shifts: "+str(shift_count))
		siteinfo.append("SHIFT A:"+a1+":00-"+b1+":00 Hours")
		siteinfo.append("SHIFT B:"+b1+":00-"+c1+":00 Hours")
		siteinfo.append("SHIFT C:"+c1+":00-"+a1+":00 Hours")


	# if sitetyp=="only night":
	# 	siteinfo.append("Monitoring happens only in Night Shift")
	# else:
	siteinfo.append("Monitoring happens 24 Hours")

	siteinfo.append("Patrol Monitoring Frequency: Every "+str(int(gap))+" hours")

	if site_code=="Emami-Pacharia":
		siteinfo.append("12 QR codes are not yet installed for this site")




	# if gap=="1hr":
	# 	siteinfo.append("Post Monitoring Frequency: Every 1 Hour")
	# elif gap=="mixed":
	# 	siteinfo.append("Post Monitoring Frequency: Every 30 minutes in Shift B. Every 1 hour in Shift A.")
	# else:
	# 	siteinfo.append("Post Monitoring Frequency: Every 30 minutes")
	


	tz = timezone.get_current_timezone()

	time1=datetime.now()
	# time1=tz.localize(time1
	time2_1=time1-timedelta(hours=time1.hour,minutes=time1.minute)
	time2_1=time2_1+timedelta(hours=float(a))
	time2_11=time2_1-timedelta(days=1)


	# endhr=24-starthr
	# time2_11=time2_1-timedelta(hours=endhr)
	# time2_1=time2_1+timedelta(hours=starthr)
	
	end_time1=tz.localize(time2_1)
	start_time1=tz.localize(time2_11)

	starthr=int(start_time1.hour)
	# print(starthr)


	# print(start_time1,end_time1)

	endtm=end_time1-timedelta(minutes=1)

	dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(endtm,"%d-%B %I:%M %p")

	print(dt_1)

	# endtm=end_time1-timedelta(minutes=1)

	# dt_1=datetime.strftime(start_time1,"%d-%B %I:%M %p") +" to "+ datetime.strftime(time1,"%d-%B %I:%M %p")

	all_beacons_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=site_code,uuidtype="QR PATROL").values()))
	# print(len(all_beacons_names))

	all_place_name=pd.DataFrame(list(placeTbl.objects.filter(siteid__name__icontains=site_code).values()))

	place_list=all_place_name['name'].tolist()


	# last30=datetime.datetime.strftime(start_time1,"%d %b %Y")+" to "+datetime.datetime.strftime(end_time1,"%d %b %Y")


	all_path=pd.DataFrame(list(pathTbl.objects.filter(uuid__placeid__siteid__name__icontains=site_code,uuid__uuidtype="QR PATROL",time__range=(start_time1,end_time1)).values()))

	all_beacons_names=all_beacons_names.rename(columns={"uuid":"uuid_id"})

	siteinfo.append("Total No. of Places: "+str(len(all_beacons_names)))




	# dict_hour={"time":[]}

	all_beacons_names["fullname"]=all_beacons_names["name"]+"\n\n"+all_beacons_names["placeid_id"]

	if placeselect=="":

		name_list=all_beacons_names["fullname"].tolist()

		nameonly_list=all_beacons_names["name"].tolist()

	else:

		name_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"fullname"].tolist()
		nameonly_list=all_beacons_names.loc[all_beacons_names["placeid_id"]==placeselect,"name"].tolist()
		# particular_beacon_names=pd.DataFrame(list(uuidTbl.objects.filter(placeid__name__icontains=placeselect).values()))

	# for ele in name_list:
	# 	dict_hour[ele]=[]

	last_visit=[]
	visitcount=[]
	visitby=[]
	colorls=[]
	attainedslot=[]
	missedslot=[]

	if len(all_path)==0:

		# last_visit=[]
		# visitcount=[]
		# visitby=[]
		# colorls=[]
		i=0
		for i in range(len(all_beacons_names)):
			last_visit.append("Not visited")
			visitcount.append(0)
			visitby.append("Nobody visited")
			colorls.append("color:#ff0000")
			attainedslot.append("None")
			missedslot.append("24 Hours")


		


	else:
		df_all1=pd.merge(all_path,all_beacons_names,on='uuid_id')

		all_emp=pd.DataFrame(list(empTbl.objects.all().values()))
		all_emp=all_emp.rename(columns={"empid":"empid_id","name":"empname"})

		df_all=pd.merge(df_all1,all_emp,on='empid_id')

		df_all['time'] = pd.to_datetime(df_all['time'], infer_datetime_format=True).dt.tz_convert('Asia/Kolkata')

		df_all["datetime"]=df_all["time"].dt.strftime('%d %b %Y %I %p')
		df_all["datetime1"]=df_all["time"].dt.strftime('%d %b %I:%M %p')
		df_all["dateonly"]=df_all["time"].dt.date
		df_all["hour"]=df_all["time"].dt.hour
		df_all["minute"]=df_all["time"].dt.minute

		# print(df_all["datetime1"])

		uuidls=all_beacons_names["uuid_id"].tolist()

		for ele in uuidls:
			df_all_1=df_all.loc[df_all["uuid_id"]==ele]

			if len(df_all_1)==0:
				last_visit.append("Not visited")
				visitcount.append(0)
				visitby.append("Nobody visited")
				colorls.append("color:#ff0000")
				attainedslot.append("None")
				missedslot.append("24 Hours")

			else:
				visitcount.append(len(df_all_1))
				# hrls=df_all_1["hour"].tolist()


				mis_slot=""
				attendslot=0

				x1=int(float(a))
				freq1=int(gap)
				k=0
				while k<tot_slot:
					if x1+gap<24:
						if len(df_all_1.loc[(df_all_1["hour"]>=x1)&(df_all_1["hour"]<x1+gap)])>0:
							attendslot=attendslot+1
						else:
							mis_slot=mis_slot+" "+str(int(x1))+"-"+str(int(x1+gap))+" hours,"
					else:
						newgap=0+(x1+gap-24)
						if len(df_all_1.loc[(df_all_1["hour"]>=x1)|(df_all_1["hour"]<newgap)])>0:
							attendslot=attendslot+1
						else:
							mis_slot=mis_slot+" "+str(int(x1))+"-"+str(int(newgap))+" hours,"



					if x1+gap>=24:
						x1=0+(x1+gap-24)
					else:
						x1=x1+gap
					k=k+1

				attainedslot.append(attendslot)
				missedslot.append(mis_slot)

				if attendslot==tot_slot:
					colorls.append("color:#00ff00")
				else:
					colorls.append("color:#fed300")











				last_time=df_all_1["datetime1"].tolist()[-1]
				last_emp=df_all_1["empname"].tolist()[-1]
				last_visit.append(last_time)
				visitby.append(last_emp)

				# time2=df_all_1["time"].tolist()[-1]

				# timegap=(end_time1-time2).total_seconds()
				# timegap=timegap/3600
				# if timegap>=gap:
				# 	colorls.append("color:#ff0000")
				# else:
				# 	colorls.append("color:#00ff00")
				# gaplist.append(timegap)


		name_list1=[]
		for ele10 in name_list:
			x1=ele10[0:(len(ele10)-len(site_code)-2)]
			name_list1.append(x1)




	df_final=pd.DataFrame({"place":name_list1,"visit":visitcount,"visitat":last_visit,"visitby":visitby,"att":attainedslot,"gap":missedslot,"colorls":colorls})

	df_final=df_final.sort_values(["gap"],ascending=False)

	# print(df_final)















		



		# start_hr=starthr
		# end_hr=starthr+1


		# i=0
		# now_hour=int(time1.hour)

		# total_slot=abs(now_hour-start_hr+1)
		# if now_hour>=20:
		# 	total_slot=abs(now_hour-start_hr+1)
		# elif now_hour<8:
		# 	total_slot=abs((24-starthr)+now_hour+1)
		# else:
		# 	total_slot=abs(now_hour-start_hr+1)
		# total_slot=abs(now_hour-start_hr+1)
		# print("this is total slot", total_slot)
		


	# df_final=pd.DataFrame(dict_hour)
	# # print(counter)
	# # print(len(dict_hour))

	# # for key in dict_hour:
	# # 	print(len(dict_hour[key]),key)

	json_records3 = df_final.reset_index().to_json(orient ='records')
	data3 = []
	data3 = json.loads(json_records3)


	# records=df_final.to_records(index=False)
	# result=tuple(records)
	# result=result[::-1]
	# # print(result)

	# name_list1=[]


	




	context={"namelist":name_list1,"dt_1":dt_1,"data":data3,"placelist":place_list,"sites":site_code,"beac":len(all_beacons_names),"starthr":starthr,"gap":gap, "siteinfo":siteinfo}


	return render(request,"livedatap.html",context)
	# return HttpResponse("Success")







