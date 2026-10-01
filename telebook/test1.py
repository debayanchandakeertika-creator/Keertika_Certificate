# x=[0, 10, 15, 20, 1, 5, 60, 8, 79, 0]


# i=0
# curren_ele=0
# prev_ele=0
# picks=[]
# troughs=[]
# for ele in x:
# 	curren_ele=ele
# 	# if curren_ele>=prev_ele:
# 	# 	picks.append(curren_ele

# 	curren_ele=ele


# 	if prev_ele>curren_ele:
# 		picks.append(prev_ele)
# 	else:
# 		troughs.append(curren_ele)



# 	prev_ele=curren_ele

# 	# if (curren_ele-prev_ele)<0:
# 	# 	picks.append(curren_ele)
# 	# 	prev_ele=curren_ele
# 	# else:
# 	# 	troughs.append(curren_ele)
# 	# 	prev_ele=curren_ele




# # import pandas as pd

# x=[{"text1":"ELECTRICAL","text2":"","dryer1":"60","dryer2":"60","dryer3":"60","dryer4":"60"},
# {"text1":"THERMAL","text2":"","dryer1":"50","dryer2":"50","dryer3":"50","dryer4":"50"},
# {"text1":"COMBINED","text2":"","dryer1":"40","dryer2":"40","dryer3":"40","dryer4":"40"}]

# y=[]
# i=0

# while i<len(x):
# 	ele=x[i]
# print(picks,troughs)



# # for ele in x:
# # 	y=ele.keys()
# # 	print(y)




# # print(y)

	







# [{"dryer":"1","thermal":"50","electrical":"60","combined":"40"},
# {"dryer":"2","thermal":"50","electrical":"60","combined":"40"},
# {"dryer":"3","thermal":"50","electrical":"60","combined":"40"},
# {"dryer":"4","thermal":"50","electrical":"60","combined":"40"}]



#  pip3 install requests

# Replace the deviceToken key with the device Token for which you want to send push notification.
# Replace serverToken key with your serverKey from Firebase Console

# Run script by ->
# python3 fcm_python.py


import requests
import json

serverToken = 'AAAA2t9jZa0:APA91bHN2njv-Da-53xd-EV32aX7yMzXxlzA4DmU4wbB1ogndwMZACr0puvXuFApWb_hw8NkEGMEellJNbwX_vTYG64_IBbkHwGtSuRFKsABczokD84ZZD8I0Fg6rR61wjSJY4wQ8m03'
deviceToken = 'fUaq0cGRRs6OfSzTpXdNlu:APA91bEh6wZmgLWvWuKaSu_ixgypCrwFzsQ_zI1BoQvfD0peVG0KSgEjmeNlDGQc4ducTPKyXlzv9g9Lls1aFFPsLZqpx1pcoMHCjBmEEJ0jPBIe94SCHMXWQ43yPsqBVlZP2gJWjDOg'

headers = {
        'Content-Type': 'application/json',
        'Authorization': 'key=' + serverToken,
      }

body = {
          'notification': {'title': 'Sending push form python script',
                            'body': 'New Message'
                            },
          'to':
              deviceToken,
          'priority': 'high',
        #   'data': dataPayLoad,
        }
response = requests.post("https://fcm.googleapis.com/fcm/send",headers = headers, data=json.dumps(body))
print(response.status_code)

print(response.json())