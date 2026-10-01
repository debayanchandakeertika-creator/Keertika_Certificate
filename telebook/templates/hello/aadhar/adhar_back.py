import os
import re
import cv2
import requests
import argparse
import pytesseract
import numpy as np
from pytesseract import Output

from pythreshold.utils import test_thresholds,apply_threshold
from pythreshold.local_th import sauvola_threshold, niblack_threshold, wolf_threshold,\
    nick_threshold, lmean_threshold, bradley_roth_threshold,\
    bernsen_threshold, contrast_threshold, singh_threshold, feng_threshold

def state_dist(pin):
    url="https://api.postalpincode.in/pincode/"+pin
    res=requests.get(url)
    state=res.text.split('State":"')[1].split('"')[0]
    dist=res.text.split('"District":"')[1].split('"')[0]
    return state,dist

def return_pin(text):
    try:
        match = re.search(r'\d{6}', text)
        a,b=match.span()
        return text[a:b]
    except Exception as e:
        return "Not Found"
        
def return_aadhar(text):
    aad="Not Found"
    encoded_string = text.encode("ascii", "ignore")
    decode_string = encoded_string.decode()
    for  i  in decode_string.split("\n"):
        ls = ''.join(re.findall(r'\d+', i))
        if len(ls)==12:
            aad=ls
    return aad

def thres_img(img):   
    img=cv2.medianBlur(img,1)
    th = feng_threshold(img)
    thres_img = apply_threshold(img, th)
    kernel = np.ones((1, 1), np.uint8)
    thres_img = cv2.morphologyEx(thres_img, cv2.MORPH_OPEN, kernel)
    thres_img = cv2.morphologyEx(thres_img, cv2.MORPH_CLOSE, kernel)
    return thres_img

def roi2(path):
    adrs=pin=state=dist=adhar_no = "Not found"
    
    #img = cv2.imread('E:/vmsared/Aadhar_images/small_card_back/6.jpeg')
    img = cv2.imread(path)
    img1 = thres_img(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    #config='--psm 6 -l eng+ben+hin -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyz0123456789'
    d = pytesseract.image_to_data(img1, output_type=Output.DICT,config='--psm 6 -l eng+ben+hin' )
    keys = list(d.keys())
    listw=[]
    #date_pattern = '^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/(19|20)\d\d$'
    date_pattern = re.compile(r'\d\d\d\d\d\d')
    
    #print(d)
    n_boxes = len(d['text'])
    x1=y1=x2=y2=0
    x=0
    y=0
    w=0
    h=0
    t=0
    t1=0
    offset=2
    
    yy1,xx1=img1.shape
    yyy1,xxx1=img1.shape
    #print(yy1,xx1)
    flag=0
    for i in range(n_boxes):
        if int(d['conf'][i]) > 60:
            (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
            flag=0
            if re.match(date_pattern, d['text'][i]):
                img = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
                t= y + 2*h
                #print('pin found')
            elif  re.match('S/O', d['text'][i]) or re.match('D/O', d['text'][i]) or re.match('W/O', d['text'][i]) or re.match('C/O', d['text'][i]):
#                 x1=x-(2*w)
#                 y1=y-(2*h)
#                 x2=x+(6*w)
#                 y2=y+(8*w)
                #print('s/o found')
                if flag==1:
                    pass
                img = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
#                 img = cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                x=x-3*w
                listw = listw + [[ x , y ]]
                #print(listw)
                
                
                #break
            elif re.match('Address', d['text'][i]):
#                 x1=x-(2*w)
#                 y1=y-(2*h)
#                 x2=x+(6*w)
#                 y2=y+(8*w)
                #print('address found')
                flag=1
                t1=y+4*h
                img = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
#                 img = cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                x=x-2*w
                listw = listw + [[ x , y ]]
                #print(listw)
                
                #break
            #listw.append([x, y])
    #print(listw)
    #print("===============================================")
    try:
        if t<10:
            t=t1
        x1=min(map(lambda x: x[0], listw))
        y1=max(map(lambda x: x[1], listw))

        img = cv2.rectangle(img, (x1, y1-10), (xx1, t), (255, 255, 255), 2)
        if x1 <1:
            x1=1
        custom_config = r'-l eng+ben+hin'
        adrs = pytesseract.image_to_string(img1[y1-10 : t, x1 : xx1-30], config=custom_config)   
        pin = return_pin(adrs)
        state,dist = state_dist(pin)
        custom_config_1='outputbase digits -c tessedit_char_whitelist=0123456789'
        number_only = pytesseract.image_to_string(img1[yyy1//2 : yyy1 , 0 : xxx1], config=custom_config_1)
        adhar_no = return_aadhar(number_only)

    finally:
        return state,dist,pin,adrs,adhar_no
    
