import pyodbc
import datetime
import os
import xlsxwriter

message=''
server_name="vrsqlecm\\vrsqlfilenet"
db="Nvision_Custom"
query=''
status=[]
now = str(datetime.date.today()).replace("-","_")
col=0
row=0
cur_wrk_dir=os.getcwd()
attachment_flag=False
to_mail_list=[]
cc_mail_list=[]

#Get the query
def get_query():
	global cur_wrk_dir
	global query
	file=open(cur_wrk_dir+"\\Utility\\query.txt","r")
	for line in file:
		query+=line

#Geth the difference between two dates
def days_between(processed_date, current_date):
    processed_date = datetime.datetime.strptime(processed_date, "%Y-%m-%d")
    current_date = datetime.datetime.strptime(current_date, "%Y-%m-%d")
    return abs((current_date - processed_date).days)

#Insert Excel Header
def excel_header():
	global worksheet
	global workbook
	global row
	global col
	worksheet.write(row,col,"NSM_Loan_number")
	col+=1
	worksheet.write(row,col,"P8_Doc_GUID")
	col+=1
	worksheet.write(row,col,"Batch_Name")
	col+=1
	worksheet.write(row,col,"Doc_Ingestion_Status")
	col+=1
	worksheet.write(row,col,"Brand_Id")
	col+=1
	worksheet.write(row,col,"Cloud_Move_Indicator")
	col+=1
	worksheet.write(row,col,"Website_Cloud_Reference_ID")
	col+=1
	worksheet.write(row,col,"Processed_Date")
	col+=1
	worksheet.write(row,col,"Web_Complete_Date")
	col+=1
	worksheet.write(row,col,"Upload_Pending_Days")
	row+=1
	col=0

#Write status to excel
def write_to_excel():
	global worksheet
	global workbook
	global row
	global col
	
	excel_header()

	for dict in status:
		for key in dict.keys():
			worksheet.write(row,col,str(dict[key]))
			col+=1
		row+=1
		col=0

#Getting the mailIds from the file
def get_mail_list(file_name,type):
	global to_mail_list
	global cc_mail_list
	file=open(file_name,"r")
	for line in file:
		if ".com" in line:
			if type == "to":
				to_mail_list.append(line)
			elif type == "cc":
				cc_mail_list.append(line)
	
#Setting up the mailing list
def set_mail_list():
	to_mail=os.getcwd()+"\\Utility\\to_address.txt"
	cc_mail=os.getcwd()+"\\Utility\\cc_address.txt"
	get_mail_list(to_mail,"to")
	get_mail_list(cc_mail,"cc")
	
#Mail status to users
def sendEmail(message_body):
	global output_file_name
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	set_mail_list()
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	mail_cc=";".join(cc_mail_list)
	mail.CC = mail_cc
	if attachment_flag == True:
		mail.Attachments.Add(output_file_name)
	mail.Subject = 'Azure Statement Upload Status - '+ now
	mail.Body = message_body
	try:
		mail.Send()
	except Exception:
		print('')
	
#Connection to database		
con = pyodbc.connect(driver="{SQL Server}",server=server_name,database=db,Trusted_Connection='yes')
cur = con.cursor()
print("-------------------")
print("Azure Upload Status")
print("-------------------\n")

#Get the query
get_query()

#Executing the query and store the result
results = cur.execute(query)
for result in results:
	dict={}
	dict['NSM_Loan_number']=result[0]
	dict['P8_Doc_GUID']=result[1]
	dict['Batch_Name']=result[2]
	dict['Doc_Ingestion_Status']=result[3]
	dict['Brand_Id']=result[4]
	dict['Cloud_Move_Indicator']=result[5]
	dict['Website_Cloud_Reference_ID']=result[6]
	dict['Processed_Date']=str(result[7])[0:10]
	dict['Web_Complete_Date']=str(result[8])[0:10]
	status.append(dict)
	del(dict)
	

output_file_name=cur_wrk_dir+"\\Archive\\Azure_Upload_Status_"+now+".xlsx"
workbook = xlsxwriter.Workbook(output_file_name)
worksheet = workbook.add_worksheet()
message+="Hi Team,"
if len(status) != 0:
	attachment_flag=True
	message+="\n\tThere are some documents which are not yet uploaded into the Azure cloud. Please Find Attached the Azure Document Upload Status."
	write_to_excel()
else:
	message+="\n\tAzure Upload process is working fine. No documents were lagged!"

message+="\n\nThanks & Regards,\nImaging Support"
if attachment_flag == True:
	sendEmail(message)
	print("Azure Upload has been mailed!\n")
else:
	print("No files were lagged in Azure upload")
workbook.close()
a=input("Press enter to quit the process")
if a != '':
	exit()
	
	
	
	
	

