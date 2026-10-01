from django.contrib import admin
from .models import *
from django.db.models import Count

# Register your models here.

# class u_display(admin.ModelAdmin):
# 	list_display =["mobile","tid","email"]

# class d_display(admin.ModelAdmin):
# 	list_display =["occupancy","name"]

# class m_display(admin.ModelAdmin):
# 	list_display =["start","end","tid","desk_name","is_cancelled"]
# 	def desk_name(self, obj):
# 		return obj.desk.name

# admin.site.register(users,u_display)
# admin.site.register(deske,d_display)
# admin.site.register(meetings,m_display)
# #admin.site.register(meetings,CategoryChoiceField)

class uuid_display(admin.ModelAdmin):

	list_display =["uuid",'location1',"other","uuidtype"]
	def location1(self,obj):
		return f"{obj.placeid.siteid.companyid.name}, {obj.placeid.siteid.name}, {obj.placeid.building}, {obj.placeid.floor}, {obj.placeid.name}"
	location1.short_description='Location(company, site, building, floor, place)'
	search_fields =['placeid__siteid__companyid__name','placeid__siteid__name','placeid__building','placeid__floor','placeid__name']
	# def get_search_results(self, request, queryset, search_term):
	# 	# search_term is what you input in admin site
	# 	search_term_list = search_term.split(' ')  #['apple','bar']
	# 	# queryset, use_distinct = super(uuid_display, self).get_search_results(request, queryset, search_term)
	# 	if not any(search_term_list):
	# 		return queryset, False

	# 	for item in search_term_list:
	# 		queryset |= self.objects.filter(placeid_id__siteid__companyid__name__icontains=item)
	# 		queryset |= self.objects.filter(placeid_id__siteid__name__icontains=item)
	# 		queryset |= self.objects.filter(placeid_id__building__icontains=item)
	# 		queryset |= self.objects.filter(placeid_id__floor__icontains=item)
	# 	print(queryset,search_term_list)
	# 	return queryset, False


class emp_display(admin.ModelAdmin):
	list_display =["empid","name","phone"]

class path_display(admin.ModelAdmin):
	list_display =["uuid","mac","dis",'rssi','time','empid']
	search_fields =['time','uuid__other','empid__name']
	list_filter =['time','uuid__other','empid__name']
	list_display_links=['uuid','empid']

class site_display(admin.ModelAdmin):

	list_display=['siteid','name','location','companyid','about']
	search_fields =['siteid','name','location','about']
	list_filter=['location']
	list_display_links=['companyid','siteid']

class place_display(admin.ModelAdmin):
	list_display=['siteid','name','floor','building','about']

class company_display(admin.ModelAdmin):
	list_display=['companyid','name','detail']

class success_display(admin.ModelAdmin):
	list_display=['siteid','date1','succrate']

class notification_display(admin.ModelAdmin):
	list_display=['empid1','firebaseid','type1']

class qs_display(admin.ModelAdmin):
	list_display=['uuid','qs','rightans','type1','type2']

class qsstore_display(admin.ModelAdmin):
	list_display=['uuid','qs','ans','type1','time','empid']

class mail_display(admin.ModelAdmin):
	list_display=["siteid","emailid1","emailid2","emailid3"]


admin.site.register(uuidTbl,uuid_display)
admin.site.register(empTbl, emp_display)
admin.site.register(pathTbl,path_display)
admin.site.register(siteTbl,site_display)
admin.site.register(placeTbl,place_display)
admin.site.register(companyTbl,company_display)
admin.site.register(successTbl,success_display)
admin.site.register(notificationTbl,notification_display)
admin.site.register(qsTbl,qs_display)
admin.site.register(qsstoreTbl,qsstore_display)
admin.site.register(mailTbl,mail_display)
