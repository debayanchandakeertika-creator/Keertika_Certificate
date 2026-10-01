# # importing the qrcode library  
# import qrcode
# import img2pdf
# from PIL import Image,ImageDraw,ImageFont
# import os
# import uuid
# from uuid import getnode
# import random
# import sys
# import time
# from datetime import datetime


# import requests
# import datetime
# current_time3=datetime.datetime.now().strftime("%d %B %Y %I:%M:%S %p")
# print(current_time3)
# x=requests.get(f"http://127.0.0.1:8000/storeData?mac=239a1f17-990d-4dd9-b1c6-eca34ae8956d&dis=0&rssi=0&time={current_time3}&empid=123")

# # print (':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff)
# # for elements in range(0,2*6,2)][::-1]))
# # typ=sys.argv[1]
# # if typ=='mac':
# # x=("%02x:%02x:%02x:%02x:%02x:%02x" % (
# #         random.randint(0, 255),
# #         random.randint(0, 255),
# #         random.randint(0, 255),
# #         random.randint(0, 255),
# #         random.randint(0, 255),
# #         random.randint(0, 255),

# #         ))
# # print(x)
# # 	# x=uuid.uuid4()
# # # else:
# # print(uuid.uuid1())
# # print(x.getnode())
# # print(type(x))
# # generating a QR code using the make() function  
# qr_img = qrcode.make("Hello World",box_size=40,border=4,)  
# # saving the image file 
# #qr_img.make_image(fill_color="red", back_color="#23dda0",text="HELLO")
# qr_img.save("qr-img.png")


# img1 = Image.open("qr-img.png")
# img2 = Image.open("blank1.png")

# # print(img2.shape())

# Image.Image.paste(img2, img1, (17, 96))
# #img2.save("qr-img.png")

# draw = ImageDraw.Draw(img2)

# draw.text((450, 30),"Sample Text",'black',font = ImageFont.truetype("arial.ttf",30))
# img2.save("qr-img.png")


# ls=['a','b','c']
# for idx, val in enumerate(ls):
#     print(idx, val)
# draw = ImageDraw.Draw(img)
# # font = ImageFont.truetype(<font-file>, <font-size>)
# #font = ImageFont.truetype("sans-serif.ttf", 16)
# # draw.text((x, y),"Sample Text",(r,g,b))
# draw.text((0, 0),"Sample Text",'red')
# img.save('sample-out.jpg')

# pdf_bytes = img2pdf.convert("qr-img.png")
# pdf_path="qr1.pdf"
# file = open(pdf_path, "wb")
# file.write(pdf_bytes)
# file.close()
  
# # output
# print("Successfully made pdf file")

import numpy as np

def detect_peaks_and_troughs(data):
    peaks_indices = []
    troughs_indices = []
    
    for i in range(1, len(data) - 1):
        if data[i] > data[i-1] and data[i] > data[i+1]:
            peaks_indices.append(data[i])
        elif data[i] < data[i-1] and data[i] < data[i+1]:
            troughs_indices.append(data[i])
    
    return peaks_indices, troughs_indices


data=[1,2,3,5,7,1,2,0,4,7,11,13,9,7]

x,y=detect_peaks_and_troughs(data)

print(x,y)
print(data[11])
