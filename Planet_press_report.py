import os
import pyodbc
import getpass
import fnmatch
import shutil
from shutil import move
import xlsxwriter
import datetime
import time

now = str(datetime.date.today()).replace("-","")
cur_wrk_dir=os.getcwd()
usaa_source_dir=''
cooper_source_dir=''
output_file_name=cur_wrk_dir+"\\Archive\\output\\PP_Status-"+now+".xlsx"
destination_dir=cur_wrk_dir+"\\Input\\"
archive_dir=cur_wrk_dir+"\\Archive\\"
workbook = xlsxwriter.Workbook(output_file_name)
worksheet = workbook.add_worksheet()
more_reasons=False
failed_files_status=[];
reason_count=0
row=0
col=0
message=''
mrcooper=False
usaa=False
to_mail_list=[]
cc_mail_list=[]

#Fetch files from the server
def fetch_files_from_server():
	global destination_dir
	global usaa_source_dir
	global cooper_source_dir
	source_path_to_exceptions=cur_wrk_dir+"\\Utility\\input_paths.txt"
	paths=[]
	exception_path_input_file=open(source_path_to_exceptions,"r")
	for line in exception_path_input_file:
		paths.append(line.split(" - ")[1])
		
	cooper_source_dir=paths[0][0:len(paths[0])-1]
	usaa_source_dir=paths[1]
	files_to_be_moved = ["mrcooper_files.txt","usaa_files.txt"]
	for file_to_be_moved in files_to_be_moved:
		if "mrcooper_files" in file_to_be_moved:
			#source_dir = "\\\\CRFNUTLPRD04\\Cooper\\Jobs\\LettersandStatmentExceptions\\"
			cooper_file_path=cur_wrk_dir+"\\Input\\"+"mrcooper_files.txt"
			cooper_file=open(cooper_file_path,"w")
			for f in os.listdir(cooper_source_dir):
				if fnmatch.fnmatch(f, "*.zip"):
					cooper_file.write(str(f))
					cooper_file.write("\n")
			cooper_file.close()
		elif "usaa_files" in file_to_be_moved:
			#source_dir = "\\\\CRFNUTLPRD02\\Jobs\\LettersandStatmentExceptions\\"
			usaa_file_path=cur_wrk_dir+"\\Input\\"+"usaa_files.txt"
			usaa_file=open(usaa_file_path,"w")
			for f in os.listdir(usaa_source_dir):
				if fnmatch.fnmatch(f, "*.txt"):
					usaa_file.write(str(f))
					usaa_file.write("\n")
			usaa_file.close()
		#move(source_dir,destination_dir)
		
#Fetching the status of the batch
def batch_status(filename,db,brand):
    global count
    #Querying the Database
    query = """SELECT [Batch_Name],
                      [Status],
                      [Total_Docs_Recvd],
                      [Total_Docs_Inserted]
               FROM ["""+db+"""].[dbo].[Statements_Letters_Batch_Status]
               WHERE Batch_Name="""+"'"+filename+"'"
    results = cur.execute(query);

    #Stripping the results
    for result in results:
        single_file_status={};
        single_file_status['S_No']=count
        single_file_status['Brand']=brand
        single_file_status['Batch_Name']=result[0]
        single_file_status['Status']=result[1]
        single_file_status['Received']=result[2]
        single_file_status['Inserted']=result[3]
        single_file_status['Reason1']=''
        single_file_status['Reason2']=''
        #del(single_file_status)

    count+=1
    del(results)
    del(query)
    return single_file_status

#Fetching the failure reason of the batches
def failure_reason(filename,db):
    query = """SELECT [Batch_Name],
                      [Failure_Reason]
               FROM ["""+db+"""].[dbo].[Statements_Letters_Failed_Loans]
               WHERE Batch_Name="""+"'"+filename+"'"
    results = cur.execute(query);

    #Filter the failure reason from the resultset
    failure_reason_list=[];
    for result in results:
        if "doesn't have a doc type name" in result[1]:
            if "Doesn't have a doc type name" not in failure_reason_list:
                failure_reason_list.append("Doesn't have a doc type name")
        else:
            failure_reason_list.append(result[1])
    return failure_reason_list

#Fetching the list of files that are failed
def fetch_file_names(filename):
    failed_files=[]
    failed_file=open(filename,"r")
    for line in failed_file:
        if ".zip" in line:
            failed_files.append(line[0:len(line)-1])
    return failed_files

#Build the failed file status
def build_status(failed_files,db,brand):
	for a in failed_files:
		global more_reasons
		current_batch_status=batch_status(str(a),db,brand)
		current_batch_failure_reasons=failure_reason(str(a),db)
		if len(current_batch_failure_reasons)==1:
			reason_count=1
			current_batch_status['Reason1']=current_batch_failure_reasons[0]
			current_batch_status['Reason2']=''
		elif len(current_batch_failure_reasons)==2:
			reason_count=2
			more_reasons=True
			current_batch_status['Reason1']=current_batch_failure_reasons[0]
			current_batch_status['Reason2']=current_batch_failure_reasons[1]
		else:
			reason_count=3
			current_batch_status['Reason1']="The Batches failed due to so many reasons!Please verify!"
		failed_files_status.append(current_batch_status)
		
#Insert Excel Header
def excel_header():
	global worksheet
	global row
	global col
	worksheet.write(row,col,"S_No")
	col+=1
	worksheet.write(row,col,"Brand")
	col+=1
	worksheet.write(row,col,"Batch")
	col+=1
	worksheet.write(row,col,"Status")
	col+=1
	worksheet.write(row,col,"Received")
	col+=1
	worksheet.write(row,col,"Ingested")
	col+=1
	if more_reasons == True:
		worksheet.write(row,col,"Reason 1")
		col+=1
		worksheet.write(row,col,"Reason 2")
	else:
		worksheet.write(row,col,"Reason")
	row+=1
	col=0

#Write status to excel
def write_to_excel():
	global worksheet
	global row
	global col
	
	excel_header()
	previous_brand=''
	current_brand=''

	for dict in failed_files_status:
		current_brand=dict['Brand']
		if (current_brand != previous_brand) and (previous_brand != ''):
			row+=1
			excel_header()
		for key in dict.keys():
			worksheet.write(row,col,str(dict[key]))
			col+=1
		row+=1
		col=0
		previous_brand=current_brand
	workbook.close()

#Generate the content for the mail body
def generate_mail_body():
	global message
	if mrcooper and usaa:
		message="Hi Team,\n\tFile(s) have failed during the Planet Press Ingestion Process under both 'Mr.cooper' and 'USAA'. PFA the list of failed files."
	elif mrcooper and not usaa:
		message="Hi Team,\n\tFile(s) have failed during the Planet Press Ingestion Process under 'Mr.cooper'. PFA the list of failed files."
	elif usaa and not mrcooper:
		message="Hi Team,\n\tFile(s) have failed during the Planet Press Ingestion Process under 'USAA'. PFA the list of failed files."
	else:
		message="Hi Team,\n\tNo file(s) have failed during the Planet Press Ingestion Process."
	message=message+"\n\nThanks & Regards,\nImaging Support";

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
	if mrcooper or usaa:
		mail.Attachments.Add(output_file_name)
	mail.Subject = 'Planet Press Statement and Exceptions - '+ now
	mail.Body = message_body
	try:
		mail.Send()
	except Exception:
		print('')

#Moving files to archive folder after mailing
def move_files_to_archive():
	mrcooper_src=destination_dir+"mrcooper_files.txt"
	mrcooper_dstn=archive_dir+"mrcooper\\mrcooper_files_"+now+".txt"
	shutil.move(mrcooper_src,mrcooper_dstn)
	usaa_src=destination_dir+"usaa_files.txt"
	usaa_dstn=archive_dir+"usaa\\usaa_files_"+now+".txt"
	shutil.move(usaa_src,usaa_dstn)

print("Processing.....\n")
fetch_files_from_server()

for f in os.listdir(destination_dir):
    if fnmatch.fnmatch(f, '*.txt'):
        if "mrcooper" in str(f):
            count=1
            brand="mrcooper"
            db="Nvision_Custom"
            failed_files=fetch_file_names(destination_dir+str(f))
            con = pyodbc.connect(driver="{SQL Server}",server="vrsqlecm\\vrsqlfilenet",database=db,Trusted_Connection='yes')
            cur = con.cursor()
            if len(failed_files)>0:
                mrcooper=True
            build_status(failed_files,db,brand)

        elif "usaa" in str(f):
            count=1
            brand="usaa"
            db="USAA_Custom"
            failed_files=fetch_file_names(destination_dir+str(f))
            con = pyodbc.connect(driver="{SQL Server}",server="vrsqlecm\\vrsqlfilenet",database=db,Trusted_Connection='yes')
            cur = con.cursor()
            if len(failed_files)>0:
                usaa=True
            build_status(failed_files,db,brand)

write_to_excel()
generate_mail_body()
sendEmail(message)
move_files_to_archive()

print("Planet Press Statements and Exception reports have been mailed!\n")

a=input("Press enter to quit the process")
if a != '':
	exit()
	

