import pyodbc
import datetime
import os

now = str(datetime.date.today())
aborted_nsm=[]
aborted_usaa=[]
to_mail_list=[]
cc_mail_list=[]

#Fetch Verify pending Count
def verify_pending(brand):
	query="""select count(*) FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
			where qu_task = 'Verify' and qu_status = 'pending' and PB_BRANDID ="""
	query=query+"'"+brand+"'"
	results = cur.execute(query)
	for result in results:
		count=result[0]
	return count
	
#Fetch Verify hold Count
def verify_hold(brand):
	query="""select count(*) FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
			where qu_task = 'Verify' and qu_status = 'hold' and PB_BRANDID ="""
	query+="'"+brand+"'"
	results = cur.execute(query)	
	for result in results:
		count=result[0]
	return count

#Fetch Verify running Count
def verify_running(brand):
	query="""select count(*) FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
			where qu_task = 'Verify' and qu_status = 'running' and PB_BRANDID ="""
	query+="'"+brand+"'"
	results = cur.execute(query)	
	for result in results:
		count=result[0]
	return count

#Fetch Fixup pending/hold Count
def fixup_pending_hold(brand):
	query="""select count(*) FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
			where qu_task = 'fixup' and qu_status in ('pending','hold') and PB_BRANDID ="""
	query+="'"+brand+"'"
	results = cur.execute(query)	
	for result in results:
		count=result[0]
	return count
	
#Fetch aborted Count
def aborted(brand):
	query="""select count(*) FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
			where qu_status = 'aborted' and PB_BRANDID ="""
	query+="'"+brand+"'"
	results = cur.execute(query)	
	for result in results:
		count=result[0]

	#Fetch aborted batches
	if count > 0:
		query="""select pb_batch FROM [ECM_Check_PRD_DCEngine].[dbo].[JobMonitor] 
				where qu_status = 'aborted' and PB_BRANDID ="""
		query+="'"+brand+"'"
		results = cur.execute(query)	
		for result in results:
			if brand == "NSM":
				aborted_nsm.append(result[0])
			elif brand == "USAA":
				aborted_usaa.append(result[0])
	return count

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
	to_mail=os.getcwd()+"\\Mailing List\\to_address.txt"
	cc_mail=os.getcwd()+"\\Mailing List\\cc_address.txt"
	get_mail_list(to_mail,"to")
	get_mail_list(cc_mail,"cc")
	
#Send the status mail
def sendEmail(resultMail):
	global to_mail_list
	global cc_mail_list
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	set_mail_list()
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	mail_cc=";".join(cc_mail_list)
	mail.CC=mail_cc
	mail_subject="Check App--Export and Pending "+str(now)
	mail.Subject = mail_subject
	mail.Body = resultMail
	try:
		mail.send()
	except Exception:
		print('')

#List of brands
brands=[
	"NSM",
	"USAA"
]
				
#Connect to database
status=[];
db="ECM_Check_PRD_DCEngine"
con = pyodbc.connect(driver="{SQL Server}",server="vrsqlimgrev\\img_rev_prod",database=db,Trusted_Connection='yes')
cur = con.cursor()

message_body="Hi All,"
#Get the count of jobs in each category
for brand in brands:
	message_body += "\n\n" + brand + ":"
	message_body += "\n\t" + str(verify_pending(brand)) + " : Batches in Verify/Pending"
	message_body += "\n\t" + str(verify_hold(brand)) + " : Batches in Verify/Hold"
	verify_running_count = verify_running(brand)
	message_body += "\n\t" + str(verify_running_count) + " : Batches in Verify/Running"
	if verify_running_count > 0:
		message_body += "(Set to pending)"
	message_body += "\n\t" + str(fixup_pending_hold(brand)) + " : Batches in FixUp/Hold or pending"
	aborted_count = aborted(brand)
	message_body += "\n\t" + str(aborted_count) + " : Batches in aborted"
	if aborted_count > 0:
		message_body += "(Set to pending)"
		message_body += "\n\t    (Aborted Batch Id - "
		if brand == "NSM":
			for single_batch in aborted_nsm:
				message_body += str(single_batch)
				aborted_count-=1
				if aborted_count != 0:
					message_body += ", "
		if brand == "USAA":
			for single_batch in aborted_usaa:
				message_body += str(single_batch)
				aborted_count-=1
				if aborted_count != 0:
					message_body += ","
		message_body += ")"
	
message_body+="\n\nThanks & Regards,\nImaging Support"
sendEmail(message_body)
print("CheckApp Export and Pending reports have been mailed!\n\nReset the batches in the Silo!\n")

a=input("Press enter to quit the process")
if a != '':
	exit()
	
