from django.db import models
from django.conf import settings
from django.core.validators import RegexValidator
# Create your models here.
# class users(models.Model):
# 	mobile = models.CharField(max_length=12,blank=False)
# 	tid=models.IntegerField(blank=False)
# 	email=models.CharField(max_length=100,blank=True)
# 	def __unicode__(self):
# 		return self.mobile

# class deske(models.Model):
# 	name=models.CharField(max_length=50,blank=False)
# 	occupancy=models.IntegerField(blank=False)
# 	def __unicode__(self):
# 		return self.name
# 	def __str__(self):
# 		return self.name

# class meetings(models.Model):
# 	start=models.DateTimeField()
# 	end=models.DateTimeField()
# 	tid=models.IntegerField(blank=False)
# 	desk=models.ForeignKey(deske,on_delete=models.CASCADE)
# 	is_cancelled=models.BooleanField(default=False)
# 	def __unicode__(self):
# 		return self.tid
class companyTbl(models.Model):
	companyid=models.CharField(primary_key=True,max_length=50)
	name=models.CharField(max_length=50)
	detail=models.CharField(default='None',max_length=50)
	def __str__(self):
		return f"{self.name}"

class siteTbl(models.Model):
	siteid=models.CharField(primary_key=True,max_length=100)
	companyid=models.ForeignKey(companyTbl, on_delete=models.CASCADE)
	name=models.CharField(max_length=50)
	location=models.CharField(default='None',max_length=50)
	about=models.CharField(default='None',max_length=50)
	def __str__(self):
		return f"{self.name}"

class placeTbl(models.Model):
	placeid=models.CharField(primary_key=True,max_length=100)
	siteid=models.ForeignKey(siteTbl, on_delete=models.CASCADE)
	name=models.CharField(max_length=100)
	floor=models.CharField(default='None',max_length=50)
	building=models.CharField(default='None',max_length=50)
	about=models.CharField(default='None',max_length=50)
	def __str__(self):
		return f"{self.name}"
class uuidTbl(models.Model):

	uuid=models.CharField(primary_key=True, max_length=50)
	placeid=models.ForeignKey(placeTbl, on_delete=models.CASCADE)
	other=models.CharField(max_length=50)
	name=models.CharField(max_length=50,blank=True)
	uuidtype=models.CharField(max_length=50,blank=True)
	time1=models.DateTimeField(blank=True)
	def __str__(self):
		return f"{self.other}, {self.uuid}"
	def location(self):
		return f"{self.placeid.siteid.companyid.name},{self.placeid.siteid.name},{self.placeid.building},{self.placeid.floor},{self.placeid.name}"

class empTbl(models.Model):

	empid=models.CharField(primary_key=True,max_length=50)
	name=models.CharField(max_length=50,blank=False)
	phone_regex = RegexValidator(regex=r'^\+?1?\d{9,15}$', message="Phone number must be entered in the format: '+999999999'. Up to 15 digits allowed.")
	phone=models.CharField(validators=[phone_regex],max_length=15,blank=True)
	employee_type=models.CharField(max_length=100,blank=True)
	def __str__(self):
		return f"{self.name}, {self.empid}"

class pathTbl(models.Model):

	uuid=models.ForeignKey(uuidTbl, on_delete=models.CASCADE)
	mac=models.CharField(max_length=50)
	dis=models.FloatField()
	lat=models.FloatField(blank=True,null=True,default=0.0)
	long=models.FloatField(blank=True,null=True,default=0.0)
	rssi=models.IntegerField(blank=False)
	time=models.DateTimeField(blank=True)
	empid=models.ForeignKey(empTbl, on_delete=models.CASCADE)


class successTbl(models.Model):
	siteid=models.ForeignKey(siteTbl, on_delete=models.CASCADE)
	date1=models.DateField(blank=True)
	succrate=models.FloatField(blank=True,null=True,default=0.0)
	nightsuccrate=models.FloatField(blank=True,null=True,default=0.0)
	gsuccrate=models.FloatField(blank=True,null=True,default=0.0)

class notificationTbl(models.Model):
	empid1=models.CharField(primary_key=True,max_length=100)
	firebaseid=models.CharField(max_length=200)
	deviceid=models.CharField(max_length=200)
	type1=models.CharField(max_length=100,blank=True,default="")
	type2=models.CharField(max_length=100,blank=True,default="")
	type3=models.CharField(max_length=100,blank=True,default="")
	desg=models.CharField(max_length=100,blank=True,default="")

class detailTbl(models.Model):
	siteid=models.ForeignKey(siteTbl, on_delete=models.CASCADE)
	status=models.CharField(max_length=100,blank=True,default="")
	A1=models.IntegerField(null=True,blank=True,default=0)
	B1=models.IntegerField(null=True,blank=True,default=0)
	C1=models.IntegerField(null=True,blank=True,default=0)
	G1=models.IntegerField(null=True,blank=True,default=0)
	G2=models.IntegerField(null=True,blank=True,default=0)
	M1=models.IntegerField(null=True,blank=True,default=0)
	M2=models.IntegerField(null=True,blank=True,default=0)
	dp=models.CharField(max_length=100,blank=True,default="")
	csd=models.CharField(max_length=100,blank=True,default="")
	night=models.IntegerField(null=True,blank=True,default=0)
	AQ=models.IntegerField(null=True,blank=True,default=0)
	BQ=models.IntegerField(null=True,blank=True,default=0)
	CQ=models.IntegerField(null=True,blank=True,default=0)
	GQ=models.IntegerField(null=True,blank=True,default=0)
	MQ=models.IntegerField(null=True,blank=True,default=0)

class qsTbl(models.Model):
	qsid=models.CharField(primary_key=True,max_length=50)
	uuid=models.ForeignKey(uuidTbl, on_delete=models.CASCADE)
	qs=models.CharField(max_length=300,blank=False)
	rightans=models.CharField(max_length=50,blank=False)
	type1=models.CharField(max_length=200,blank=True,default="")
	type2=models.CharField(max_length=100,blank=True,default="")
	type3=models.CharField(max_length=100,blank=True,default="")
	type4=models.CharField(max_length=50,blank=True,default="")
	type5=models.CharField(max_length=100,blank=True,default="")


class qsstoreTbl(models.Model):
	uuid=models.ForeignKey(uuidTbl, on_delete=models.CASCADE)
	empid=models.ForeignKey(empTbl, on_delete=models.CASCADE)
	qs=models.CharField(max_length=300,blank=False)
	ans=models.CharField(max_length=50,blank=False)
	type1=models.CharField(max_length=200,blank=True,default="")
	type2=models.CharField(max_length=100,blank=True,default="")
	type3=models.CharField(max_length=100,blank=True,default="")
	time=models.DateTimeField(blank=True)

class mailTbl(models.Model):
	siteid=models.ForeignKey(siteTbl, on_delete=models.CASCADE)
	emailid1=models.CharField(max_length=60,blank=True,default="")
	emailid2=models.CharField(max_length=60,blank=True,default="")
	emailid3=models.CharField(max_length=60,blank=True,default="")
	emailid4=models.CharField(max_length=60,blank=True,default="")
	emailid5=models.CharField(max_length=60,blank=True,default="")
	emailid6=models.CharField(max_length=60,blank=True,default="")
	emailid7=models.CharField(max_length=60,blank=True,default="")
	emailid8=models.CharField(max_length=60,blank=True,default="")
	emailid9=models.CharField(max_length=60,blank=True,default="")
	emailid10=models.CharField(max_length=60,blank=True,default="")
