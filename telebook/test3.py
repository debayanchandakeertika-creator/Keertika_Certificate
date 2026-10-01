
import requests
import json
import pandas as pd


headers={"AuthUsername":"Achilles","AuthToken":"rSODFoSpEOMOeHBJXHlMDpgdCOGZJVxrrVdGbjvBlxvOmCBrrS"}

url="http://outsidevendorservice.starspayroll.in:700/Payroll/RestService_Payroll.svc/Employee_Payroll_in_Particularmonth?cmpid=1&region=&branchcode=&datesearchfield=&datefrom=&dateto=&categorylevel=&category_code=&emptype=&emp_status=&employeeid=&branchgroup=&deptcode=&salary_Month=11&salary_Year=2022"


a = requests.get(url, headers=headers)

# print(a)

data=json.loads(a.text)

df=pd.json_normalize(data["Employee_Payroll_in_ParticularmonthResult"])
# df.to_csv("allsalary.csv")


df1=df.loc[(df["Branchcode"]=="FCM-BBR")|(df["Branchcode"]=="SEC-BBR")]

df1=df1.reset_index()

empid_ls=df1["EmployeeID"].tolist()

name1=[]
mobile_no1=[]
designation1=[]
branch_ls=[]


for empid in empid_ls:

	if empid!="" and empid[0]!="0":
		empid="0"+empid

	url1=f"http://outsidevendorservice.starspayroll.in:700/CommonServices_OutsideModule/Service_CommonServices.svc/Get_Employee?empid={empid}&emptype=%&empstatus=%&regionid=%&branchid=%&companyid=%&groupcompanyid=%&divisioncode=%"

	try:
		a1=requests.get(url1)
		print(a1)
	except Exception as e:
		print(e)
	data1=json.loads(a1.text)

	if len(data1["Get_EmployeeResult"])!=0:
		empname=data1["Get_EmployeeResult"][0]["employeename"]
		phone_num=data1["Get_EmployeeResult"][0]["cmobileno"]
		desg=data1["Get_EmployeeResult"][0]["designation"]
		branch1=data1["Get_EmployeeResult"][0]["Branchcode"]

		name1.append(empname)
		mobile_no1.append(phone_num)
		designation1.append(desg)
		branch_ls.append(branch1)



df_final=pd.DataFrame({"name":name1,"branch":branch_ls,"designation":designation1,"mobile no.":mobile_no1})


df_final.to_csv("BBR_emp.csv")
