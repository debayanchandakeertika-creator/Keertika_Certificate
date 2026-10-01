# x=[7,8,9,9,2,5,5,3,5,2,1]
# #Result -> [1,3,7,8,2,2,5,5,9,9]

# dict_1={}

# for ele in x:
# 	if ele not in dict_1.keys():
# 		dict_1[ele]=1

# 	else:
# 		dict_1[ele]=dict_1[ele]+1

# # print(dict_1)

# output_ls=[]
# min_freq=0

# # dict_2=dict_1.sort()
# # print(dict_2)

# # val_list=dict_1.values()

# # print(val_list)
# reverse_dict={}

# for key,values in dict_1.items():

# 	if values not in reverse_dict.keys():
# 		reverse_dict[values]=[]
# 		reverse_dict[values].append(key)
# 	else:
# 		reverse_dict[values].append(key)

# # print(reverse_dict)

# for keys in reverse_dict:
# 	y=reverse_dict[keys]
# 	y.sort()
# 	# print(y)
	
# 	for ele2 in y:
# 		for i in range(keys):
# 			output_ls.append(ele2)
# 	# print(keys)

# print(output_ls)

# Python3 implementation of the approach
 
# Function that returns true if num is
# greater than both arr[i] and arr[j]
# def isPeak(arr, n, num, i, j):
 
#     # If num is smaller than the element
#     # on the left (if exists)
#     if (i >= 0 and arr[i] > num):
#         return False
 
#     # If num is smaller than the element
#     # on the right (if exists)
#     if (j < n and arr[j] > num):
#         return False
#     return True
 
# # Function that returns true if num is
# # smaller than both arr[i] and arr[j]
# def isTrough(arr, n, num, i, j):
 
#     # If num is greater than the element
#     # on the left (if exists)
#     if (i >= 0 and arr[i] < num):
#         return False
 
#     # If num is greater than the element
#     # on the right (if exists)
#     if (j < n and arr[j] < num):
#         return False
#     return True
 
# def printPeaksTroughs(arr, n):
 
#     print("Peaks : ", end = "")
 
#     # For every element
#     for i in range(n):
 
#         # If the current element is a peak
#         if (isPeak(arr, n, arr[i], i - 1, i + 1)):
#             print(arr[i], end = " ")
#     print()
 
#     print("Troughs : ", end = "")
 
#     # For every element
#     for i in range(n):
 
#         # If the current element is a trough
#         if (isTrough(arr, n, arr[i], i - 1, i + 1)):
#             print(arr[i], end = " ")
 
# # Driver code
# arr = [5, 10, 5, 7, 4, 3, 5]
# n = len(arr)
 
# printPeaksTroughs(arr, n)

def find_longest_palindrome(s):
    longest = ''
    n = len(s)
    for i in range(n):
        for j in range(i+1,n+1):
            word = s[i:j]
            if word == word[::-1]:
                if len(word)>len(longest):
                    longest = word          
    return longest

print(find_longest_palindrome('locoannamadam'))



# print(output_ls)


	# print(key,values)