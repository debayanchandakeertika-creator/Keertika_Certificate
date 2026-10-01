import cv2
from PIL import Image
import pytesseract
from pytesseract import Output
import re
from datetime import datetime
import os
import requests
from termcolor import colored, cprint
from pythreshold.utils import test_thresholds,apply_threshold
from pythreshold.local_th import sauvola_threshold, niblack_threshold, wolf_threshold,\
    nick_threshold, lmean_threshold, bradley_roth_threshold,\
    bernsen_threshold, contrast_threshold, singh_threshold, feng_threshold

def state_dist(pin):
    try:
        url="https://api.postalpincode.in/pincode/"+pin
        res=requests.get(url)
        state=res.text.split('State":"')[1].split('"')[0]
        dist=res.text.split('"District":"')[1].split('"')[0]
        #print(res.text)
        return state,dist
    except Exception as e:
        a=["Not Found","Not Found"]
        return a
    
def return_pin(text):
    try:
        match = re.search(r'\d{6}', text)
        a,b=match.span()
        return text[a:b]
    except Exception as e:
        return "Not Found"

def clean_up(text):
    encoded_string = text.encode("ascii", "ignore")
    decode_string = encoded_string.decode()
    return decode_string

def ocr_core(filename,lng):
    text = pytesseract.image_to_string(Image.open(filename),lang=lng)
    return text

def crop_front(path):
    imagePath = path
    
    image = cv2.imread(imagePath)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    faceCascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = faceCascade.detectMultiScale(gray,scaleFactor=1.3,minNeighbors=3,minSize=(30, 30))
    hh,ww=h_w(path)
    if len(faces)>0:
        for (x, y, w, h) in faces:
            #cv2.rectangle(image, (x, y-h), (x + w*5, y + h*3), (0, 255, 0), 2)
            crop_img=image[y-int(h*0.8):y+int(h*3.2),x:x+int(w*4.8)]
            #print(hh,ww,x,y)
            address_crop_img=image[0:y-int(h*4),0:ww]
        #"/Users/suman/Desktop/open_cv/hello/static/AnonymousUser1586551554.jpeg"
        np=path.split(path.split('/')[-1])[0]+"extract/face_"+path.split('/')[-1]
        np2=path.split(path.split('/')[-1])[0]+"extract/address_"+path.split('/')[-1]
        #print(np)
        status = cv2.imwrite(np,crop_img)
        status = cv2.imwrite(np2,address_crop_img)

def thres_img(img):
    th = feng_threshold(img)
    thres_img = apply_threshold(img, th)
    return thres_img

def roi(path):
    #img = cv2.imread('E:/vmsared/Aadhar_images/small_card_back/6.jpeg')
    img = cv2.imread(path)
    hh,ww=h_w(path)
    img1 = thres_img(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    d = pytesseract.image_to_data(img1, output_type=Output.DICT,config='--psm 6 ' )
    keys = list(d.keys())
    n_boxes = len(d['text'])
    for i in range(n_boxes):
        if int(d['conf'][i]) > 60:
            if (re.match("^To$", d['text'][i]) or re.match("^To:$", d['text'][i]) ):
                (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
                #img = cv2.rectangle(img, (x-3, y), (x + (ww-x-w), y + (14*h)), (0, 255, 0), 2)
                #cv2.rectangle(image, (x, y-h), (x + w*5, y + h*3), (0, 255, 0), 2)
                #img=image[y-int(h*0.8):y+int(h*3.2),x:x+int(w*4.8)]
                img=img[y-int(h-3):y+int(h*18),x:x+int(ww-x-w)]
    img_path=path.split(path.split('/')[-1])[0]+"extract/"+path.split("/")[-1]
    #print(path)
    #print(img_path)
    #print("================")
    cv2.imwrite(img_path, img)


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
            return ('Front_Page_with_more_information')
        else:
            return "Find_Back_Page"
    else:
        return "Back_Page"


# exit()

# a=os.listdir(os.getcwd())


# for i in a:
#     if(i[-4:]=="jpeg"):
#         #print (find_face(i) + " : " + i)
#         classify=find_face(i)
#         #print(classify)
#         if classify=="Front_Page_with_more_information":
#             continue
#             #text=ocr_core(i,"eng+ben")
#             #roi(i)
#             crop_front(i)
#             roi("extract/"+"address_"+i)
#             text=ocr_core("extract/"+"address_"+i,"eng")
#             print(return_pin(text))
#             pin_code=return_pin(text)
#             if pin_code!="Not Found":
#                 print (state_dist(return_pin(text)))
#             print(i)
#             print("=======================")
#         elif(classify=="Back_Page"):
#             import adhar_back
#             try:
#                 print(adhar_back.roi2(i))
#             except Exception as e:
#                 pass
#         elif(classify=="Find_Back_Page"):
#             pass