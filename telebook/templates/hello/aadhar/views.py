from django.shortcuts import render
from django.contrib.auth import logout
from django.contrib.auth.models import User
from hello.settings import *
from django.http import HttpResponse
import hashlib
import random
import cv2
import sys
import os
from PIL import Image
import pytesseract
import re
from datetime import datetime
from termcolor import colored, cprint
import aadhar.adhar_back as adhar_back
import aadhar.rotate_image as fft
from unicodedata import normalize
from django.views.decorators.csrf import csrf_exempt
import json

def return_pin(text):
	try:
		match = re.search(r'\d{6}', text)
		a,b=match.span()
		return text[a:b]
	except Exception as e:
		return "Not Found"

def clean_up(text):
	encoded_string = text.encode("ascii", "ignore")
	decode_string = encoded_string.decode(unicode_escape)
	return decode_string

@csrf_exempt
def api(request):

	msg=""
	a="%s" %(request.user)
	if request.method == 'POST':

		up=upload_file(request.FILES['test'],a)
		msg="<img src=static/"+up+">"
		img_url="static/"+up
		path=os.path.dirname(BASE_DIR)+"/hello/static/"+up
		#rotate_img(path)
		classify=find_face(path)
		#print(classify)
		if classify=="Find_Back_Page":
			crop_front(path)
			text=(ocr_core(path,"eng"))
			find_aadhar=return_aadhar(ocr_core(path,"eng+ben"))
			dob=return_dob(text)
			father=return_father(text)
			gender=return_gender(text)
			name=return_name(ocr_core(path,"eng+ben"))

			jdata={
			"name":name,
			"father":father,
			"dob":dob,
			"gender":gender,
			"aadhar":find_aadhar,
			"img":img_url,
			}

			msg=json.dumps(jdata)

			#msg+="<br><br>Name:&nbsp"+name+"<br>Father:&nbsp"+father+"<br>DOB:&nbsp"+dob+"<br>Gender:&nbsp"+gender+"<br>Aadhar:&nbsp;"+find_aadhar
		elif classify=="Back_Page":
			back_details=adhar_back.roi2(path)
			state=back_details[0]
			dist=back_details[1]
			pin=back_details[2]
			#adrs=back_details[3].replace("\n","<br>")
			adrs=back_details[3].encode('ascii', 'ignore').decode('unicode_escape')
			adhar_no=back_details[4]
			jdata={
			"address":adrs,
			"state":state,
			"district":dist,
			"pincode":pin,
			"aadhar":adhar_no,
			"img":img_url,
			}
			msg=json.dumps(jdata)
			#msg+="<br><br>"+adrs+"<br>State:&nbsp"+state+"<br>District:&nbsp"+dist+"<br>Pin:&nbsp"+str(pin)+"<br>Aadhar:&nbsp"+adhar_no
		elif classify=="Front_Page_with_more_information":
			#print (path)
			fft.crop_front(path)
			addr_path=path.split(path.split('/')[-1])[0]+"extract/address_"+path.split('/')[-1]
			face_path=path.split(path.split('/')[-1])[0]+"extract/face_"+path.split('/')[-1]
			fft.roi(addr_path)
			texted=ocr_core(addr_path,"eng+ben")
			find_pin_text=ocr_core(addr_path,"eng")
			text=(ocr_core(face_path,"eng"))
			find_aadhar=return_aadhar(ocr_core(face_path,"eng+ben"))
			dob=return_dob(text)
			father=return_father(text)
			gender=return_gender(text)
			name=return_name(ocr_core(face_path,"eng+ben"))
			#msg+="<br><br>Name:&nbsp"+name+"<br>Father:&nbsp"+father+"<br>DOB:&nbsp"+dob+"<br>Gender:&nbsp"+gender+"<br>Aadhar:&nbsp;"+find_aadhar

			texted=normalize('NFKD',texted)
			text2=""
			for i in texted.split("\n"):
				encoded_string = i.encode("ascii", "ignore")
				i = encoded_string.decode().strip()
				if i!='':	
					text2+=i+"\n"
			pin=fft.return_pin(find_pin_text)
			
			try:
				text2=text2.split("To")[1].split(pin)[0]+pin
			except Exception as e:
				pass
			
			#print(pin)
			sd=fft.state_dist(pin)
			kk="rm -rf "+addr_path
			ll="rm -rf "+face_path
			os.system(kk)
			os.system(ll)

			jdata={
			"name":name,
			"father":father,
			"dob":dob,
			"gender":gender,
			"aadhar":find_aadhar,
			"address":text2,
			"state":sd[0],
			"district":sd[1],
			"pincode":pin,
			"img":img_url,
			}
			msg=json.dumps(jdata)
			#msg+="<br><br>+"+text2+"<br>"+"<br>State:&nbsp"+sd[0]+"<br>District:&nbsp"+sd[1]

	context={
	"msg":msg,
	}
	return render(request,'api.html',context)

@csrf_exempt
def index(request):

	msg=""
	a="%s" %(request.user)
	if request.method == 'POST':

		up=upload_file(request.FILES['test'],a)
		msg="<img src=static/"+up+">"
		path=os.path.dirname(BASE_DIR)+"/hello/static/"+up
		#rotate_img(path)
		classify=find_face(path)
		#print(classify)
		if classify=="Find_Back_Page":
			crop_front(path)
			text=(ocr_core(path,"eng"))
			find_aadhar=return_aadhar(ocr_core(path,"eng+ben"))
			dob=return_dob(text)
			father=return_father(text)
			gender=return_gender(text)
			name=return_name(ocr_core(path,"eng+ben"))
			msg+="<br><br>Name:&nbsp"+name+"<br>Father:&nbsp"+father+"<br>DOB:&nbsp"+dob+"<br>Gender:&nbsp"+gender+"<br>Aadhar:&nbsp;"+find_aadhar
		elif classify=="Back_Page":
			back_details=adhar_back.roi2(path)
			state=back_details[0]
			dist=back_details[1]
			pin=back_details[2]
			adrs=back_details[3].replace("\n","<br>")
			adhar_no=back_details[4]
			msg+="<br><br>"+adrs+"<br>State:&nbsp"+state+"<br>District:&nbsp"+dist+"<br>Pin:&nbsp"+str(pin)+"<br>Aadhar:&nbsp"+adhar_no
		elif classify=="Front_Page_with_more_information":
			#print (path)
			fft.crop_front(path)
			addr_path=path.split(path.split('/')[-1])[0]+"extract/address_"+path.split('/')[-1]
			face_path=path.split(path.split('/')[-1])[0]+"extract/face_"+path.split('/')[-1]
			fft.roi(addr_path)
			texted=ocr_core(addr_path,"eng+ben")
			find_pin_text=ocr_core(addr_path,"eng")
			text=(ocr_core(face_path,"eng"))
			find_aadhar=return_aadhar(ocr_core(face_path,"eng+ben"))
			dob=return_dob(text)
			father=return_father(text)
			gender=return_gender(text)
			name=return_name(ocr_core(face_path,"eng+ben"))
			msg+="<br><br>Name:&nbsp"+name+"<br>Father:&nbsp"+father+"<br>DOB:&nbsp"+dob+"<br>Gender:&nbsp"+gender+"<br>Aadhar:&nbsp;"+find_aadhar

			texted=normalize('NFKD',texted)
			text2=""
			for i in texted.split("\n"):
				encoded_string = i.encode("ascii", "ignore")
				i = encoded_string.decode().strip()
				if i!='':	
					text2+=i+"<br>"
			pin=fft.return_pin(find_pin_text)
			
			try:
				text2=text2.split("To")[1].split(pin)[0]+pin
			except Exception as e:
				pass
			
			#print(pin)
			sd=fft.state_dist(pin)
			kk="rm -rf "+addr_path
			ll="rm -rf "+face_path
			os.system(kk)
			os.system(ll)
			msg+="<br><br>+"+text2+"<br>"+"<br>State:&nbsp"+sd[0]+"<br>District:&nbsp"+sd[1]

	context={
	"msg":msg,
	}
	return render(request,'render.html',context)

def ocr_core(filename,lng):
	pytesseract.pytesseract.tesseract_cmd='C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
	text = pytesseract.image_to_string(Image.open(filename),lang=lng)
	return text

def return_name(text):
	try:
		name="File Format not supported"
		match = re.search(r'\d{2}/\d{2}/\d{4}', text)
		a,b=match.span()
		bottom_up=[]
		for i in text[:a].split("\n"):
			encoded_string = i.encode("ascii", "ignore")
			i = encoded_string.decode().strip()
			if len(i)>3:
				bottom_up.append(i)
		bottom_up=bottom_up[::-1]
		if 'father' in bottom_up[1].lower():
			name=bottom_up[2]
		else:
			name=bottom_up[1]
		name = (re.sub(r'\W +', '',name))
		return name
	except Exception as e:
		return "Not Found"


def return_gender(text):
	if "female" in text.lower():
		return  "Female"
	else:
		return "Male"

def return_aadhar(text):
	aad="Not Found"
	encoded_string = text.encode("ascii", "ignore")
	decode_string = encoded_string.decode()
	for  i  in decode_string.split("\n"):
		ls = ''.join(re.findall(r'\d+', i))
		if len(ls)==12:
			aad=ls
	return aad

def return_father(text):
	ls="Not Found"
	for i in text.split("\n"):
		if("father" in  i.lower()):
			ls = (re.sub(r'\W +', '', i.lower().split("father")[1]).strip())
	return ls

def return_dob(text):
	try:
		match = re.search(r'\d{2}/\d{2}/\d{4}', text)
		a,b=match.span()
		return text[a:b]
	except Exception as e:
		return "Not Found"

def get_h_w(file):
    img = cv2.imread(file)
    height, width, channels = img.shape
    return (height/width)

def h_w(file):
    img = cv2.imread(file)
    height, width, channels = img.shape
    return (height,width)

def find_face(path):
    imagePath = path

    image = cv2.imread(imagePath)
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    faceCascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = faceCascade.detectMultiScale(gray,scaleFactor=1.3,minNeighbors=3,minSize=(30, 30))

    if len(faces)>0:
        if get_h_w(imagePath)>1:
            return 'Front_Page_with_more_information'
        else:
            return "Find_Back_Page"
    else:
        return "Back_Page"

def crop_front(path):
    imagePath = path

    image = cv2.imread(imagePath)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    faceCascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = faceCascade.detectMultiScale(gray,scaleFactor=1.3,minNeighbors=3,minSize=(30, 30))

    if len(faces)>0:
        for (x, y, w, h) in faces:
            #cv2.rectangle(image, (x, y-h), (x + w*5, y + h*3), (0, 255, 0), 2)
            crop_img=image[y-int(h*0.8):y+int(h*3.2),x:x+int(w*4.8)]
        np=imagePath
        status = cv2.imwrite(np,crop_img)
    
    return imagePath

def rotate_img(path):
    imagePath=path
    img = cv2.imread(imagePath)
    #print("=============")
    #print(find_face(img))
    if find_face_2(img)!=1:
        for x in range(1,4):
            print(find_face_2(img))
            if find_face_2(img)==1:
                break
            else:
                img=cv2.transpose(img)
                img=cv2.flip(img,flipCode=1)
                cv2.imwrite(path, img)
    #print("============")

def find_face_2(path):

    image = path
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    faceCascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = faceCascade.detectMultiScale(gray,scaleFactor=1.3,minNeighbors=3,minSize=(30, 30))
    #status = cv2.imwrite('faces_detected.jpg', image)

    return len(faces)

def upload_file(f,a):
	filename, ext = os.path.splitext(f.name)
	import time
	#return(dir(f))
	siz=f.size
	mime=f.content_type
	if (mime == 'image/png' or mime == 'image/jpeg' or mime ==  'application/pdf' or mime=='application/pdf'):
		#print mime
		mime = True
	else:
		mime = False
	#print "mime is "+str(mime)	
	if (ext == '.png' or ext == '.pdf' or ext == '.jpg' or ext == '.jpeg'):
		exti=True
	else:
		exti=False
	#print "ext is "+str(ext)
	if (siz<=20971520 and exti== True and mime == True):
		kk=str(a)+str(int(time.time()))+ext
		path=os.path.dirname(BASE_DIR)+"/hello/static/"+kk
		#print path
		destination=open(path, 'wb+')
		for chunk in f.chunks():
			#print len(chunk)
			destination.write(chunk)
		destination.close()
		if ext == '.pdf':
			comm="/usr/bin/pdfimages -j "+os.path.dirname(BASE_DIR)+"/hello/static/"+kk+" "+os.path.dirname(BASE_DIR)+"/hello/static/"+kk.split(kk.split('/')[-1])[0]+kk.split('/')[-1].split('.')[0]
			os.system(comm)
			from PyPDF2 import PdfFileReader
			pdf=PdfFileReader(os.path.dirname(BASE_DIR)+"/hello/static/"+kk)
			
			if(pdf.getDocumentInfo().author=='CamScanner'):
				one=h_w(os.path.dirname(BASE_DIR)+"/hello/static/"+kk.split('./')[-1].split('.')[0]+"-001.jpg")
				two=h_w(os.path.dirname(BASE_DIR)+"/hello/static/"+kk.split('./')[-1].split('.')[0]+"-000.jpg")
				print(one[0],two[0])
				if(one[0]>two[0]):
					return kk.split('/')[-1].split('.')[0]+"-001.jpg"
				else:
					return kk.split('/')[-1].split('.')[0]+"-000.jpg"
			#print (comm)
			#comm="rm -rf "
			else:
				return kk.split('/')[-1].split('.')[0]+"-000.jpg"
		return kk
		#return "/Users/suman/Desktop/open_cv/hello/AnonymousUser1586560173.jpeg"
		#return kk.split(kk.split('/')[-1])[0]+kk.split('/')[-1].split('.')[0]+"-000.ppm"
	else:
		return "not"
