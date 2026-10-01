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
	else:
		abc= uuidTbl(uuid=uuid_code,placeid=placeid,name=place,other=shft,uuidtype='QR',time1=datetime.now()).save()


	qr_img = qrcode.make(uuid_code,box_size=30,border=4,)
	qr_img.save(f"QR_CODES/{uuid_code}.png")

	img1 = Image.open(f"QR_CODES/{uuid_code}.png")
	img2 = Image.open("blank1.png")

	Image.Image.paste(img2, img1, (17, 132))
	#img2.save("qr-img.png")

	draw = ImageDraw.Draw(img2)

	draw.text((400, 20),"NIS Management Limited",'black',font = ImageFont.truetype("arial.ttf",45))


	draw.text((470, 68),floor+": "+place,'black',font = ImageFont.truetype("arial.ttf",25))
	draw.text((420, 98),building+ ", Site: "+site,'black',font = ImageFont.truetype("arial.ttf",25))
	if typ=="patrol":
		draw.text((520,125 ),"(Patrol Monitoring)",'black',font = ImageFont.truetype("arial.ttf",20))
	else:
		draw.text((520,125 ),"(Post Monitoring)",'black',font = ImageFont.truetype("arial.ttf",20))


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
			time=datetime.strptime(time, '%d %B %Y %I:%M:%S %p').strftime('%Y-%m-%d %H:%M')
		
		empid=request.GET.get("empid","")
		
		if abs(int(rssi))<100:
			empData=empTbl.objects.get(empid=empid)
			abc= pathTbl(uuid=uuidTbl.objects.get(uuid=mac),mac=mac,dis=dis,rssi=rssi,time=time,empid=empData,lat=lat,long=long1).save()
			success={"abc":abc}
		else:
			success={"abc":'Failed'}
	
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
					print("data submission done!")
					# os.remove(a)
			except Exception as e:
				pass

		success={"abc":'Success'}
	
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

		print(rows['f'])
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
			otal_slot=abs(now_hour-start_hr+1)

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



def qrReportTwoShift(request):
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
	dt_1=yest.strftime("%d-%B-%Y")
	# print(todayStr[-1], type(todayStr[-1]))
	# print(yestStr[-1], type(yestStr[-1]))

	# print("-----------Yesterday------->>>", yestStr, type(yestStr))
	# print("-----------Today----------->>>", todayStr, type(todayStr))
	
	# shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
	# 		(12,00,13,00), (13,00,14,00), (14,00,15,00), (15,00,16,00),
	# 		(16,00,17,00), (17,00,18,00), (18,00,19,00), (19,00,20,00)]

	shiftB = [(20,00,21,00), (21,00,22,00), (22,00,23,00), (23,00,00,00),
				(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),
				(4,00,5,00), (5,00,6,00), (6,00,7,00), (7,00,8,00)]

	# isDataShiftA = []
	# AttendedShiftA = []
	# unAttendedShiftA = []
	tz=timezone.get_current_timezone()

	# for qrid in uuidList:
	# 	isQrData = []
	# 	for t in shiftA:
	# 		start = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[0]),int(t[1]))
	# 		end = datetime(int(yestStr[0]),int(yestStr[-1]),int(yestStr[1]),int(t[2]),int(t[3]))
	# 		start=tz.localize(start)
	# 		end=tz.localize(end)
	# 		data = pathTbl.objects.filter(time__range=(start, end), uuid=qrid)
	# 		if data:
	# 			isQrData.append(1)
	# 		else:
	# 			isQrData.append(0)
	# 	unAttendedShiftA.append(isQrData.count(0))
	# 	AttendedShiftA.append(isQrData.count(1))
	# 	isDataShiftA.append(isQrData)
	# print("uuidList--->>>", uuidList, len(uuidList))
	# print("isDataShiftA----->>>>>", isDataShiftA, "\n")

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

	# maxAShiftA = max(AttendedShiftA)
	# maxUShiftA = max(unAttendedShiftA)
	maxAShiftB = max(AttendedShiftB)
	maxUShiftB = max(unAttendedShiftB)
	
	maxLen = int(max(lenList))
	module_dir = os.path.dirname(__file__)
	filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift (1).xlsx")
	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
	wb2 = xl.load_workbook(filename)
	# ws2 = wb2.active
	# ws1 = wb2["Shift A"]
	ws2 = wb2["Night Shift"]
	
	# ws1['A4'] = siteid
	ws2['A4'] = siteid
	
	# ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
	ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
	thin = Side (border_style="thin", color="000000")

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


def qrReportbhu(request):
	siteid = request.GET.get('sitecode',"")
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

	elif siteid=="AMRI-Bhubaneswar":
		shiftA = [(6,00,7,00),(7,00,8,00),(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),
			(12,00,13,00), (13,00,14,00)]

		shiftB = [(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00), (21,00,22,00)]

		shiftC = [(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00)]

		shiftG = [(8,00,9,00),(9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),
		(17,00,18,00),(18,00,19,00),(19,00,20,00),(20,00,21,00)]


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
	
	NA_shiftA_amingaon = ['Back Side DG Area', 'Loading Point']
	NA_shiftA_abhoypur = ['Factory Front side', 'Back side Utility area']

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

	post_shiftA_ag = [(7,00,7,30), (7,30,8,00), (8,00,8,30), (8,30,9,00),
				(9,00,9,30), (9,30,10,00), (10,00,10,30), (10,30,11,00), (11,00,11,30),
				(11,30,12,00), (12,00,12,30), (12,30,13,00), (13,00,13,30), (13,30,14,00),
				(14,00,14,30), (14,30,15,00), (15,00,15,30), (15,30,16,00), (16,00,16,30),
				(16,30,17,00), (17,00,17,30), (17,30,18,00), (18,00,18,30), (18,30,19,00)]

	post_shiftB_ag = [(19,00,19,30), (19,30,20,00), (20,00,20,30), (20,30,21,00),
				(21,00,21,30), (21,30,22,00), (22,00,22,30), (22,30,23,00), (23,00,23,30),
				(23,30,00,00), (00,00,00,30), (00,30,1,00), (1,00,1,30), (1,30,2,00), 
				(2,00,2,30), (2,30,3,00), (3,00,3,30), (3,30,4,00), (4,00,4,30), (4,30,5,00),
				(5,00,5,30), (5,30,6,00), (6,00,6,30), (6,30,7,00)]
	
	
	patrol_abhoypur_shiftA = [(6,00,8,00), (8,00,10,00), (10,00,12,00), (12,00,14,00), (14,00,16,00), (16,00,18,00)]
	patrol_abhoypur_shiftB = [(18,00,20,00), (20,00,22,00), (22,00,00,00), (00,00,2,00), (2,00,4,00), (4,00,6,00)]
	
	patrol_amingaon_shiftA = [(7,00,9,00), (9,00,11,00), (11,00,13,00), (13,00,15,00), (15,00,17,00), (17,00,19,00)]
	patrol_amingaon_shiftB = [(19,00,21,00), (21,00,23,00), (23,00,1,00), (1,00,3,00), (3,00,5,00), (5,00,7,00)]

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
			print(qrid[-1])
			if qrid[1] == 'QR':
				if qrid[-1] not in ('Factory Front Side', 'Back Side Utility Area'):
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
				if qrid[2] not in ('Back Side DG Area','Loading Point'):
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
					if t == (23,00,1,00):
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
					elif patrol_amingaon_shiftB.index(t) > patrol_amingaon_shiftB.index((23,00,1,00)):
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
		wsPatrolDayShift['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 06:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours"
		wsPatrolNightShift['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 18:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
	
	finalPost = []
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
			if i not in ('Factory Front Side [Area-Outside(Emami-Abhoypur)]', 'Back Side Utility Area [Area-Outside(Emami-Abhoypur)]'):
			# if i not in ('Factory Front Side', 'Back Side Utility Area'):
			# if 'Factory Front Side' or 'Back Side Utility Area' not in i:
				finalPost.append(i)
	elif siteid == 'Emami-Amingaon':
		for i in AllPost:
			# if i not in ('Back Side DG Area','Loading Point'):
			if i not in ('Back Side DG Area [Area-Outside(Emami-Amingaon)]','Loading Point [Area-Outside(Emami-Amingaon)]'):
			# if 'Back Side DG Area' or 'Loading Point' not in i:
				finalPost.append(i)
	else:
		for i in AllPost:
			finalPost.append(i)

	thin = Side (border_style="thin", color="000000")

	shiftAcolDay = 3
	shiftAiday = 0
	
	for post in finalPost:
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
	for post in allPost:
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

	col = 3
	for post in range(0, len(allPost)):
		wsPostNightShift.cell(row = 18, column = col, value=AllPost[post])
		wsPostNightShift.cell(row = 18, column = col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		wsPostNightShift.cell(row = 18, column = col).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPostNightShift.cell(row = 18, column = col).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")
		wsPostNightShift.column_dimensions[get_column_letter(col)].width = 25
		wsPostNightShift.row_dimensions[18].height = 70
		col += 1
	
	colDay = 3
	for post in range(0, len(finalPost)):
		wsPostDayShift.cell(row = 18, column = colDay, value=finalPost[post])
		wsPostDayShift.cell(row = 18, column = colDay).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
		wsPostDayShift.cell(row = 18, column = colDay).border = Border(top=thin, left=thin, right=thin, bottom=thin)
		wsPostDayShift.cell(row = 18, column = colDay).fill = PatternFill(fgColor = 'FFECB3', fill_type="solid")

		wsPostDayShift.column_dimensions[get_column_letter(colDay)].width = 25
		wsPostDayShift.row_dimensions[18].height = 70
		colDay += 1
	
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
	wsPostDayShift['E8'] = allPost[PostunAttendedShiftA.index(PostmaxUShiftA)]
	wsPostDayShift['E9'] = allPost[PostAttendedShiftA.index(PostmaxAShiftA)]
	wsPostDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPostNightShift['E8'] = allPost[PostunAttendedShiftB.index(PostmaxUShiftB)]
	wsPostNightShift['E9'] = allPost[PostAttendedShiftB.index(PostmaxAShiftB)]
	wsPostNightShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPostNightShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolDayShift['E8'] = allPatrol[PatrolunAttendedShiftA.index(PatrolmaxUShiftA)]
	wsPatrolDayShift['E9'] = allPatrol[PatrolAttendedShiftA.index(PatrolmaxAShiftA)]
	wsPatrolDayShift['E8'].alignment = Alignment(horizontal='center', vertical='center')
	wsPatrolDayShift['E9'].alignment = Alignment(horizontal='center', vertical='center')

	wsPatrolNightShift['E8'] = allPatrol[PatrolunAttendedShiftB.index(PatrolmaxUShiftB)]
	wsPatrolNightShift['E9'] = allPatrol[PatrolAttendedShiftB.index(PatrolmaxAShiftB)]
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


def qrReportbhu2s(request):
	siteid = request.GET.get('sitecode',"")
	activeQR = placeTbl.objects.filter(siteid_id=siteid)
	place_list = [i.name for i in activeQR]

	if siteid=="E-CITY Mall":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]

		shiftB = [(20,00,21,00), (21,00,22,00),(22,00,23,00), (23,00,00,00),(00,00,1,00), (1,00,2,00), (2,00,3,00), (3,00,4,00),(4,00,5,00), (5,00,6,00),(6,00,7,00),(7,00,8,00)]


	elif siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd.":
		shiftA = [(8,00,9,00), (9,00,10,00), (10,00,11,00), (11,00,12,00),(12,00,13,00), (13,00,14,00),(14,00,15,00),(15,00,16,00),(16,00,17,00),(17,00,18,00),(18,00,19,00),(19,00,20,00)]
		shiftB=[(20,00,20,30),(20,30,21,00), (21,00,21,30), (21,30,22,00), (22,00,22,30),(22,30,23,00), (23,00,23,30), (23,30,00,00), (00,00,00,30), (00,30,1,00),(1,00,1,30), (1,30,2,00), (2,00,2,30), (2,30,3,00), (3,00,3,30), 
				(3,30,4,00), (4,00,4,30), (4,30,5,00), (5,00,5,30), (5,30,6,00),(6,00,6,30),(6,30,7,00),(7,00,7,30),(7,30,8,00)]

		# print(len(shiftB))



		


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


	if siteid=="E-CITY Mall":

		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))

	elif siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd.":
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(placeid__siteid__name__icontains=siteid).values()))

	else:
		df_A=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="A"),placeid__siteid__name__icontains=siteid).values()))
		df_B=pd.DataFrame(list(uuidTbl.objects.filter(Q(other="AB")|Q(other="B"),placeid__siteid__name__icontains=siteid).values()))


		



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


	if siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd.":
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
	if siteid=="Switz Foods" or siteid=="Pearson Drums and Barrels" or siteid=="Himadri Chemical Ltd.":
		filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift_difftime.xlsx")

	else:

		filename = os.path.join(module_dir,"Post Monitoring Report Format Two Shift.xlsx")
	# BASE_DIR = Path(__file__).resolve().parent.parent
	# filename = (os.path.join(BASE_DIR, 'Post Monitoring Report Format Two Shift.xlsx'))
	wb1 = xl.load_workbook(filename)
	  
	# opening the destination excel file 
	filename1 =siteid+" Post Monitoring Report.xlsx"
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

	if siteid=="Ashwini Hospital":
	
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"
	else:
		ws1['A5'] = "Shift A :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 08:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours"
		ws2['A5'] = "Shift B :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 20:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 08:00 Hours"
		# ws3['A5'] = "Shift C :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 22:00 Hours to "+todayStr[1]+"/"+todayStr[2]+"/"+todayStr[0]+" 06:00 Hours"
		# ws4['A5'] = "Shift G :- From "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 09:00 Hours to "+yestStr[1]+"/"+yestStr[2]+"/"+yestStr[0]+" 17:00 Hours"


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
	wb2.save(str(filename1))



	if siteid=="Teloijan-Tea":
		sitenamemail="Poddar House(Teloijan-Tea)"
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