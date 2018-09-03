import xlsxwriter 
import datetime
import os
import urllib.request
from os.path import basename
import ssl
from dateutil.parser import parse
import json
import openpyxl
import shutil
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from bs4.builder import HTML
import logging
from configparser import SafeConfigParser

#Globals
logging.basicConfig(filename='status.log',level=logging.DEBUG, format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p')
cur_wrk_dir=os.getcwd()
parser = SafeConfigParser()
health_url = ''
log_urls = []
data = ''
msg_body = ''
status_report=[]
context = ssl._create_unverified_context()
log_date = ''
excel_name=''
archive_path = cur_wrk_dir+"\\Archive\\"
pass_fail_tag = True
report_table = []
to_mail_list = []
cc_mail_list = []
archive_excel_name = ''

#Step 11 - Function - Send the status mail
def sendEmail(resultMail,attachment):
	global to_mail_list
	global cc_mail_list
	global pass_fail_tag
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	mail_cc=";".join(cc_mail_list)
	mail.CC = mail_cc
	sub="Disbursement Process Monitoring "+str(log_date)
	if pass_fail_tag:
		sub = "PASS - " + sub
	else:
		sub = "FAIL - " + sub
	mail_subject=sub
	mail.Subject = mail_subject
	mail.HTMLBody = resultMail
	mail.Attachments.Add(attachment)
	try:
		mail.send()
	except Exception:
		print('')

#Step 9 - Function - Generating the table data
def generate_status_report():
	global report_table
	count=0
	for a in status_report:
		if count > 4:
			break
		report_table.append(a)
		count+=1

#Step 8 - Function - Declaring Pass or Fail Tag
def check_pass_fail_tag():
	global pass_fail_tag
	for single_report in status_report:
		if single_report['Mandatory'] == "Yes" or single_report['Mandatory'] == "yes" or single_report['Mandatory'] == "Y" or single_report['Mandatory'] == "y":
			if single_report['File_Created']!='Yes':
				pass_fail_tag = False

#Step 7 - Function - Archive report excel
def archive_report():
	global excel_name
	global archive_excel_name
	if not os.path.exists(archive_path):
		os.makedirs(archive_path)
	archive_excel_name=archive_path+excel_name
	shutil.move(excel_name,archive_excel_name)

#Step 6 - Function - Write to Excel
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
	col+=1
	worksheet.write(row,col,"Mandatory")
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
	color_format_1=workbook.add_format({'bg_color': '#E3F2FD'})
	color_format_2=workbook.add_format({'bg_color': '#BBDEFB'})
	title_color_format=workbook.add_format({'bg_color': '#2196F3'})
	merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter'})
	
	#worksheet.merge_range('B4:D4', 'Merged Range')
	#worksheet.merge_range('A2:A4', 'DISBURSEMENTS',merge_format)
	#worksheet.merge_range('A5:A10', 'SECURITIZATIONS',merge_format)
	#worksheet.merge_range('A11:A17', 'PRIVATE INVESTORS',merge_format)
	#worksheet.merge_range('A18:A23', 'CUSTOM PRIVATE INVESTORS',merge_format)
	#worksheet.merge_range('B2:B4', 'All Investors',merge_format)
	
	worksheet.conditional_format( 'A1:K1' , { 'type' : 'no_blanks' , 'format' : title_color_format} )	
	worksheet.conditional_format( 'A2:K26' , { 'type' : 'no_blanks' , 'format' : border_format} )	
	
	for single_row in range(2,row):
		if single_row % 2 == 0:
			
			worksheet.conditional_format( 'A'+str(single_row)+':K'+str(single_row) , { 'type' : 'no_blanks' , 'format' : color_format_1} )
		else:
			worksheet.conditional_format( 'A'+str(single_row)+':K'+str(single_row) , { 'type' : 'no_blanks' , 'format' : color_format_2} )
			worksheet.conditional_format( 'B'+str(single_row)+':B'+str(single_row) , { 'type' : 'blanks' , 'format' : color_format_2} )	

		
	
	worksheet.set_column(0, 0, 30)
	worksheet.set_column(1, 1, 12)
	worksheet.set_column(2, 2, 38)
	worksheet.set_column(3, 3, 27)
	worksheet.set_column(4, 10, 12)
	
	workbook.close()

#Step 5 - Function - Process Log URLs
def process_log(url):
	global context
	global status_report
	
	preprocessed_log = []
	rs_found = False
	report_name_found = False
	request=urllib.request.urlopen(url, context=context)
	response=request.read()
	data=str(response)
	data_list=data.split('\\n')
	
	for line in data_list:
		if log_date in line:
			preprocessed_log.append(line)
			
	count=1
	for single_report in status_report:
		single_report_name = single_report['Daily_Report_Name']
		for single_log in preprocessed_log:
			if "resultset size" in single_log.lower() and ("schedulerFactoryBean_Worker-").lower() in single_log.lower():
				rs=single_log.split("ResultSet Size = ")[1]
				rs_found=True
			if single_report_name[0:len(single_report_name)-8].lower() in single_log.lower() and ("schedulerFactoryBean_Worker-").lower() in single_log.lower():
				single_report['File_Created'] = 'Yes'
				single_report['Time_Created'] = single_log[11:16]
				if rs_found:
					single_report['Report_Count'] = rs
					single_report['Issue_Note'] = 'NA'
					single_report['Resolution_Note'] = 'NA' 
			if single_report['Mandatory'] == 'N' or single_report['Mandatory'] == 'n' or single_report['Mandatory'] == 'No' or single_report['Mandatory'] == 'no' or single_report['Mandatory'] == 'NO' or single_report['Mandatory'] == 'nO': 
				single_report['Issue_Note'] = 'NA'
				single_report['Resolution_Note'] = 'NA' 			
	preprocessed_log.clear()	
	
#Step 4 - Function - Reading the format file for fetching the necessary daily report names
def fetch_format_file_report_names():
	global status_report
	global non_mandate_report
	wb = openpyxl.load_workbook("Format.xlsx")
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
		singleReport['Mandatory'] = row[11].value
		status_report.append(singleReport)
		del singleReport

#Step 3 - Function- Builing the health status message body
def build_message():
	global msg_body
	global data
	total=round(int(data['diskSpace']['total'])/1073741824,4)
	free=round(int(data['diskSpace']['free'])/1073741824,4)
	msg_body="<br>Health Status"
	msg_body+="<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Overall Health&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "
	msg_body+=data['status']
	msg_body+="<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mail Health&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "
	msg_body+=data['mail']['status']
	msg_body+=" (" + data['mail']['location']+")"
	msg_body+="<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Database Health&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "
	msg_body+=data['db']['status']
	msg_body+=" (" + data['db']['database']+")"
	msg_body+="<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Disk Space: "
	msg_body+="<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Status&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "
	msg_body+=data['diskSpace']['status']
	msg_body+="<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Size&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: "
	msg_body+=str(free)+"/"+str(total)+" GB"
	if free < 2:
		msg_body+="(Space is very low!)"

#Step 2 - Function - Process Health URL
def process_health_url():
	global data
	global context
	health_response = urllib.request.urlopen(health_url, context=context)
	json_health_response = health_response.read()
	data = json.loads(json_health_response)
	
#Step 1 - Definition - Fetching data from config file
def fetch_global_declarations():
	global health_url
	global log_urls
	global to_mail_list
	global cc_mail_list
	global log_date
	parser = SafeConfigParser()
	parser.read(cur_wrk_dir + "\\config.ini")
	health_url = parser.get('links','health')
	log_urls = parser.get('links','log_urls').split(';')
	to_mail_list = parser.get('mail_ids','to').split(';')
	cc_mail_list = parser.get('mail_ids','cc').split(';')
	input_date = parser.get('log_date','process_date')
	if input_date == "Today" or input_date == "today":
		log_date = str(datetime.date.today())
	else:
		log_date = input_date

#Step 1 - Call - Fetch Data from Config File
print("Process Starting")
logging.info("Starting the process")
fetch_global_declarations()

#Step 2 - Call - Process Health URL
logging.info("Hitting the health page")
process_health_url()
logging.info("Got success response from health page")

#Step 3 - Call- Builing the health status message body
logging.info("Building health page data")
build_message()

#Step 4 - Call - Reading the format file for fetching the necessary daily report names
logging.info("Successfully fetched the links from config")
fetch_format_file_report_names()

#Step 5 - Call - Process Log URLs
logging.info("Started processing log file URLs")
for single_url in log_urls:
	logging.info("Started processing link " + single_url)
	if single_url == '':
		continue
	process_log(single_url)
	logging.info("Processing the link " + single_url + " completed!")
	
logging.info("Processing the logfile links completed")
#Step 6 - Call - Write to Excel	
excel_name="Report_"+str(log_date).replace("-","_")+".xlsx"
logging.info("Writing data to excel")
write_to_excel()

#Step 7 - Call - Archive the report
logging.info("Creating an archive for the report data")
archive_report()

#Step 8 - Call - Declaring Pass or Fail Tag
logging.info("Verifying the reports are received")
check_pass_fail_tag()

#Step 9 - Call - Generating the table data
logging.info("Finalizing the report")
generate_status_report()

#Step 10 - Creating the Mail table
logging.info("Building HTML data for mail")
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
		if key == 'Daily_Report_Name':
			table_html+="<td>"+str(a[key])+"</td>"
		else:	
			table_html+="<td align="'center'">"+str(a[key])+"</td>"
	table_html+="</tr>"

table_html+="</table>"
table_message_body='<html><body>Hi Team,<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find the Disbursement Process Status below,<br><br>'+table_html+'<p>'+msg_body+'</p><br>Thanks,<br>Chip Offshore</body></html>'

#Step 11 - Call - Mailing the Status
logging.info("Sending the mail")
sendEmail(table_message_body,archive_excel_name)
logging.info("Mailed the report successfully")
logging.info("Processing completed successfully!")
print("\nReports have been mailed!\n")