import xlsxwriter
import datetime
import os
import urllib.request
import ssl
from dateutil.parser import parse
import json
import openpyxl
import shutil
from bs4.builder import HTML

#Declarations 
context = ssl._create_unverified_context()
msg_body=''
mail_list=[]
data=''	
count=0
log_2_date = str(datetime.date.today())
#log_2_date = '2017-10-18'
status_report=[]
cur_wrk_dir=os.getcwd()
log_file_1=[]
log_file_2=[]
log_file=[]
report_table=[]
excel_name="Report_"+str(log_2_date).replace("-","_")+".xlsx"
make_archive_folder=cur_wrk_dir+"\\Archive\\"+log_2_date
if not os.path.exists(make_archive_folder):
	os.makedirs(make_archive_folder)
archive_excel_name=''

def generate_status_report():
	global report_table
	count=0
	for a in status_report:
		if count >= 3:
			break
		report_table.append(a)
		count+=1

#Step 15 - Getting the mailIds from the file
def get_mail_list(file_name):
	mail_list=[]
	file=open(file_name,"r")
	for line in file:
		if ".com" in line:	
			mail_list.append(line)
	return mail_list

#Step 14 - Fetch the mailing list
def set_mail_list(category,type):
	if category == "complete":
		if type == "to":
			mail_path=os.getcwd()+"\\Utility\\report_status_to_address.txt"
			mail_list=get_mail_list(mail_path)
			return mail_list
		if type == "cc":
			mail_path=os.getcwd()+"\\Utility\\report_status_cc_address.txt"
			mail_list=get_mail_list(mail_path)
			return mail_list
	if category == "status":
		if type == "to":
			mail_path=os.getcwd()+"\\Utility\\report_complete_to_address.txt"
			mail_list=get_mail_list(mail_path)
			return mail_list
		if type == "cc":
			mail_path=os.getcwd()+"\\Utility\\report_complete_cc_address.txt"
			mail_list=get_mail_list(mail_path)
			return mail_list
	
#Step 13 - Send the status mail
def sendEmail(resultMail,attachment):
	to_mail_list=[]
	cc_mail_list=[]
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	to_mail_list=set_mail_list("complete","to")
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	to_mail_list=set_mail_list("complete","cc")
	mail_cc=";".join(to_mail_list)
	mail.CC = mail_cc
	sub="Disbursement Process Monitoring "+str(log_2_date)
	if count == 25:
		sub = "PASS - " + sub
	else:
		sub = "FAIL - " + sub
	mail_subject=sub
	mail.Subject = mail_subject
	mail.Body = resultMail
	mail.Attachments.Add(attachment)
	try:
		mail.send()
	except Exception:
		print('')

#Step 14 - Send the status mail		
def sendTableMail(table_message_body):
	import win32com.client as win32
	to_mail_list=[]
	cc_mail_list=[]
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	to_mail_list=set_mail_list("status","to")
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	to_mail_list=set_mail_list("status","cc")
	mail_cc=";".join(to_mail_list)
	mail.CC = mail_cc
	sub="Disbursement Process Status "+str(log_2_date)
	if count == 25:
		sub = "PASS - " + sub
	else:
		sub = "FAIL - " + sub
	mail_subject=sub
	mail.Subject = mail_subject
	mail.HTMLBody = table_message_body
	try:
		mail.send()
	except Exception:
		print('')

#Step 10 - Merging the two List of dictionaries
def merge():
	global count
	global log_file_1
	global log_file_2
	for a in log_file_1:
		if a['flag'] == "1":
			count+=1
			log_file.append(a)
	for a in log_file_2:
		if a['flag'] == "1":
			count+=1
			log_file.append(a)			
	
#Step 11 - Write the result to excel
def write_to_excel():
	global excel_name
	workbook = xlsxwriter.Workbook(excel_name)
	worksheet = workbook.add_worksheet('logs')
	row = 0
	col = 0
	
	worksheet.write(row,col,"Section")
	col+=1
	worksheet.write(row,col,"Pool Name")
	col+=1
	worksheet.write(row,col,"Daily Report Name")
	col+=1
	worksheet.write(row,col,"Escalation Contact")
	col+=1
	worksheet.write(row,col,"Runtime")
	col+=1
	worksheet.write(row,col,"File_Created")
	col+=1
	worksheet.write(row,col,"Time_Created")
	col+=1
	worksheet.write(row,col,"LOB_Count")
	col+=1
	worksheet.write(row,col,"Report_Count")
	col+=1
	worksheet.write(row,col,"Issue_Note")
	col+=1
	worksheet.write(row,col,"Resolution_Note")
	col=0
	row+=1
	for dict in status_report:
		for key in dict.keys():
			worksheet.write(row,col,dict[key])
			col+=1
		row+=1
		col=0	
	
	#Formatting the Excel
	border_format = workbook.add_format()
	border_format.set_right(7)
	color_format_1=workbook.add_format({'bg_color': '#00E5FF'})
	color_format_2=workbook.add_format({'bg_color': '#00BFA5'})
	title_color_format=workbook.add_format({'bg_color': '#00BFA5'})
	merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter'})
	
	#worksheet.merge_range('B4:D4', 'Merged Range')
	worksheet.merge_range('A2:A4', 'DISBURSEMENTS',merge_format)
	worksheet.merge_range('A5:A10', 'SECURITIZATIONS',merge_format)
	worksheet.merge_range('A11:A17', 'PRIVATE INVESTORS',merge_format)
	worksheet.merge_range('A18:A23', 'CUSTOM PRIVATE INVESTORS',merge_format)
	worksheet.merge_range('B2:B4', 'All Investors',merge_format)
	
	worksheet.conditional_format( 'A1:K1' , { 'type' : 'no_blanks' , 'format' : title_color_format} )					
	worksheet.conditional_format( 'A2:K26' , { 'type' : 'no_blanks' , 'format' : border_format} )					
	worksheet.conditional_format( 'A2:K4' , { 'type' : 'no_blanks' , 'format' : color_format_1} )					
	worksheet.conditional_format( 'A5:K10' , { 'type' : 'no_blanks' , 'format' : color_format_2} )					
	worksheet.conditional_format( 'A11:K17' , { 'type' : 'no_blanks' , 'format' : color_format_1} )					
	worksheet.conditional_format( 'A18:K23' , { 'type' : 'no_blanks' , 'format' : color_format_2} )					
	worksheet.conditional_format( 'A24:K24' , { 'type' : 'no_blanks' , 'format' : color_format_1} )					
	worksheet.conditional_format( 'A25:K25' , { 'type' : 'no_blanks' , 'format' : color_format_2} )					
	worksheet.conditional_format( 'A26:K26' , { 'type' : 'no_blanks' , 'format' : color_format_1} )	
	worksheet.conditional_format( 'B26:B26' , { 'type' : 'blanks' , 'format' : color_format_1} )	

	
	worksheet.set_column(0, 0, 30)
	worksheet.set_column(1, 1, 12)
	worksheet.set_column(2, 2, 38)
	worksheet.set_column(3, 3, 27)
	worksheet.set_column(4, 10, 12)
	
	workbook.close()

#Step 12 - Archiving the log File
def archive_log(filename,code):
	global log_file_1
	global log_file_2
	file=open(filename,"w")
	if code == "1":
		log_file=log_file_1
	elif code == "2":
		log_file=log_file_2
		
	for a in log_file:
		data=str(a['s_no'])+","+str(a['result_set'])+","+str(a['daily_report_name'])+","+str(a['status'])+"\n"
		file.write(data)

#step 9 - Archive a report
def archive_report():
	global archive_excel_name
	global excel_name
	global make_archive_folder
	log1_filename=make_archive_folder+"\\Log_file_1.csv"
	archive_log(log1_filename,"1")
	log2_filename=make_archive_folder+"\\Log_file_2.csv"
	archive_log(log2_filename,"2")
	archive_excel_name=cur_wrk_dir+"\\Archive\\"+log_2_date+"\\"+excel_name
	shutil.move(excel_name,archive_excel_name)
	
#Step 8 - Generating the report
def generate_report():
	global status_report
	global log_file
	global count
	global log_file_1
	global log_file_2
	for single_report in status_report:
		daily_report_name=single_report['Daily_Report_Name']
		for a in log_file_1:
			if daily_report_name[0:len(daily_report_name)-8] == str(a['daily_report_name'])[0:len(str(a['daily_report_name']))-8]:
				if a['flag'] == "1":
					if a['status'] == "Successful":
						single_report['Report_Count']=a['result_set']
						single_report['Issue_Note']='NA'
						single_report['File_Created']='Yes'
						single_report['Resolution_Note']='NA'
						single_report['Time_Created']=a['time']
						count+=1
		for a in log_file_2:
			if daily_report_name[0:len(daily_report_name)-8] == str(a['daily_report_name'])[0:len(str(a['daily_report_name']))-8]:
				if a['flag'] == "1":
					if a['status'] == "Successful":
						single_report['Report_Count']=a['result_set']
						single_report['Issue_Note']='NA'
						single_report['File_Created']='Yes'
						single_report['Resolution_Note']='NA'
						single_report['Time_Created']=a['time']
						count+=1

#Step 7 - Processing the given URL
def process_log(url,code):
	global context
	global status_report
	global log_2_date
	global log_file_1
	global log_file_2
	
	preprocessed_log=[]
	rs_found=False
	report_name_found=False
	request=urllib.request.urlopen(url, context=context)
	response=request.read()
	data=str(response)
	data_list=data.split('\\n')

	for line in data_list:
		if log_2_date in line:
			preprocessed_log.append(line)
			
	daily_report_names=[]
	for single_report in status_report:
		daily_report_names.append(single_report['Daily_Report_Name'])
	
	count=1
	for single_report_name in daily_report_names:
		dict={}
		exception=False
		daily_report_name=''
		dict['s_no']=count
		for single_log in preprocessed_log :
			if "ResultSet Size" in single_log and ("schedulerFactoryBean_Worker-" in single_log or "schedulerFactoryBean_Worker-" in single_log):
				rs=single_log.split("ResultSet Size = ")[1]
				rs_found=True
			if single_report_name[0:len(single_report_name)-8] in single_log and ("schedulerFactoryBean_Worker-" in single_log or "schedulerFactoryBean_Worker-" in single_log):
				start_index=single_log.index(single_report_name[0:len(single_report_name)-8])
				end_index=start_index+len(single_report_name)
				daily_report_name=single_log[start_index:end_index]
				dict['result_set']=rs
				dict['time']=single_log[11:16]
				report_name_found=True
				if "successfully " in single_log.lower():
					exception=False
				else:
					exception = True
					
		if rs_found == True and report_name_found == True:
			dict['daily_report_name']=daily_report_name
			dict['status']="Successful"
			dict['flag']="1"
		elif rs_found == False and report_name_found == True:
			dict['result_set']='0'
			dict['daily_report_name']=single_report_name
			dict['status']='Resultset size not found'
			dict['flag']="0"
		else:
			dict['result_set']='0'
			dict['daily_report_name']=single_report_name
			dict['status']='Report data not found'
			dict['flag']="0"
			
		if exception == True:
			dict['status']="Exception Found! Please verify"
		
		
		if code == "1":
			log_file_1.append(dict)
		elif code == "2":
			log_file_2.append(dict)
		rs_found=False
		report_name_found=False
		del(dict)
		count+=1

#Step 6 - Reading the format file for fetching the necessary daily report names
def fetch_format_file_report_names():
	global status_report
	#Open format excel file
	format_file_path=cur_wrk_dir+"\\Utility\\Format.xlsx"
	wb = openpyxl.load_workbook(format_file_path)
	sheet=wb.active
	
	#Read the data from format excel file
	header=0
	for row in sheet:
		if header==0:
			header+=1
			continue
		singleReport={};
		singleReport['Section']=row[0].value
		singleReport['Pool_Name']=row[1].value
		singleReport['Daily_Report_Name']=row[2].value
		singleReport['Escalation_Contact']=row[3].value
		singleReport['Runtime']=row[4].value
		singleReport['File_Created']='No'
		singleReport['Time_Created']='00:00'
		singleReport['LOB_Count']=row[7].value
		singleReport['Report_Count']='0'
		singleReport['Issue_Note']='File Not Created/Excecption Found'
		singleReport['Resolution_Note']='Notify POC about the issue'
		status_report.append(singleReport)
		del singleReport

#Step 4 - Builing the health status message body
def build_message():
	global msg_body
	global data
	total=round(int(data['diskSpace']['total'])/1073741824,4)
	free=round(int(data['diskSpace']['free'])/1073741824,4)
	msg_body="Hi Team,\n\tPlease find the overall health status for this week\n"
	msg_body+="\n\t\tOverall Health      : "
	msg_body+=data['status']
	msg_body+="\n\t\tMail Health          : "
	msg_body+=data['mail']['status']
	msg_body+=" (" + data['mail']['location']+")"
	msg_body+="\n\t\tDatabase Health  : "
	msg_body+=data['db']['status']
	msg_body+=" (" + data['db']['database']+")"
	msg_body+="\n\t\tDisk Space : "
	msg_body+="\n\t\t\tStatus : "
	msg_body+=data['diskSpace']['status']
	msg_body+="\n\t\t\tSize     : "
	msg_body+=str(free)+"/"+str(total)+" GB"
	if free < 2:
		msg_body+="(Space is very low!)"
	msg_body+="\n\nThanks & Regards,\nChip Offshore"

#Step 0 - Starting Point of Execution
print("---------------------------------")
print(" Disbursement Process Monitoring")
print("---------------------------------")
print("Processing......")

#Step 1 - Fetching the URLs from the input file
urls=[]
url_file_path=os.getcwd()+"\\Utility\\links.txt"
file=open(url_file_path,"r")
for line in file:	
	link=line.split("-")[1]
	urls.append(link)
	
#Step 2 - Hit the health URL and fetch the JSON response
health_response=urllib.request.urlopen(urls[0], context=context)
json_health_response=health_response.read()

data  = json.loads(json_health_response)

#Step 3 - Process the data fetched from health page
build_message()
del(data)

#Step 5 - Process the Log File URLs
fetch_format_file_report_names()
process_log(urls[1],"1")
process_log(urls[2],"2")
#merge()
generate_report()
write_to_excel()
archive_report()
sendEmail(msg_body,archive_excel_name)
generate_status_report()

table_html=''
table_html+="<table border = "'1'">"
for a in report_table:
	table_html+="<tr bgcolor="'#2196F3'">"
	for key in a.keys():
		table_html+="<td>"+key+"</td>"
	table_html+="</tr>"
	break
for a in report_table:
	table_html+="<tr>"
	for key in a.keys():
		table_html+="<td>"+str(a[key])+"</td>"
	table_html+="</tr>"

table_html+="</table>"
table_message_body='<html><body>Hi Team,<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find the Disbursement Process Status below,<br><br>'+table_html+'<br><br>Thanks,<br>Chip Offshore</body></html>'
sendTableMail(table_message_body)
		

print("\nReports have been mailed!\n")
a=input("Press enter to quit the process")
if a != '':
	exit()
