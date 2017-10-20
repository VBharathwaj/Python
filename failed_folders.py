import os.path,shutil
import datetime,os
from shutil import move
import fnmatch
import time

#global variables to fetch the current datetime and the message_body mail string
now = str(datetime.date.today())
message_body = ''
lps_nsm_files=[]
lps_usaa_files=[]
to_mail_list=[]
cc_mail_list=[]
paths=[]

#Fetching the failed 
def get_folder_paths():
	global paths
	folder_paths_dir=os.getcwd()+"\\Utility\\paths.txt"
	paths_file=open(folder_paths_dir,"r")
	for line in paths_file:
		dir={}
		val=''
		if '\n' in line:
			val=line[0:len(line)-1]
		else:
			val=line
		if val == '':
			continue
		dir['Folder']=val.split("-")[0]
		dir['Path']=val.split("-")[1]
		paths.append(dir)
		del(dir)
		del(val)
	
#Get path for the given folder
def get_path(folder):
	path=''
	for a in paths:
		if folder.lower() in a['Folder'].lower():
			path=a['Path']
	return path

#function to get the count of failed folders
def find_count(path):
    count = len([f for f in os.listdir(path)if os.path.isfile(os.path.join(path, f))])
    return count

#Walz Failed Folders 
def walz():
	global message_body
	#path = "\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\WALZ\\Failed"
	path=get_path("walz")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tWalz................ : " + str(count) + " Failed Files"
	else:
		message_body +="\n\t\tPath to Walz failed folder is invalid! Please correct!"
		print("Path to Walz failed folder is invalid! Please correct!")
	
#Venture NSM Failed Folders
def venture_nsm():
	global message_body
	#path = "\\\\Chec.local\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\Venture\Failed"
	path=get_path("venture_nsm")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tVenture NSM.. : " + str(count) + " Failed Files"
	else:
		message_body +="\n\t\tPath to Venture NSM failed folder is invalid! Please correct!"
		print("Path to Venture NSM failed folder is invalid! Please correct!")
	
#Venture USAA Failed Folders
def venture_usaa():
	global message_body
	#path = "\\\\Chec.local\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\VentureUSAA\\Failed"
	path=get_path("venture_usaa")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tVenture USAA. : " + str(count) + " Failed Files"
	else:
		message_body +="\n\t\tPath to Venture USAA failed folder is invalid! Please correct!"
		print("Path to Venture USAA failed folder is invalid! Please correct!")
	
#Blitz Failed Folders
def blitz():	
	global message_body
	#path = "\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\BlitzDocs\\Failed"
	path=get_path("blitz")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tBlitz................ : " + str(count) + " Failed Files"
	else:
		message_body +="\n\t\tPath to Blitz failed folder is invalid! Please correct!"
		print("Path to Blitz failed folder is invalid! Please correct!")

#LPS Failed folders resolution
def lps_resolution(path,brand):
	for f in os.listdir(path):
		if fnmatch.fnmatch(f, '*.zip'):
			if brand == "nsm":
				lps_nsm_files.append(str(f))
			if brand == "usaa":
				lps_usaa_files.append(str(f))
		
#LPS NSM Failed Folders
def lps_nsm():
	global message_body
	#path = "\\\\Vrisi01\\shared$\\Rshare\CHECIT\\FTPLive\\LPSImport\\lpsimport_failed_NSM"
	path=get_path("lps_nsm")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tLPS NSM......... : " + str(count) + " Failed Files "
		if count > 0:
			lps_resolution(path,"nsm")
			message_body += "\n\t\t("
			for a in lps_nsm_files:
				message_body += a
				count -= 1
				if count != 0:
					message_body += ", "
			message_body += " - Need to be uploaded manually)"
	else:
		message_body +="\n\t\tPath to LPS NSM failed folder is invalid! Please correct!"
		print("Path to LPS NSM failed folder is invalid! Please correct!")
		
#LPS USAA Failed Folders
def lps_usaa():
	global message_body
	#path = "\\\\Vrisi01\\shared$\\Rshare\CHECIT\\FTPLive\\LPSImport\\lpsimport_failed_USAA"
	path=get_path("lps_usaa")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tLPS USAA........ : " + str(count) + " Failed Files"
		if count > 0:
			lps_resolution(path,"usaa")
			message_body += "\n\t\t("
			for a in lps_usaa_files:
				message_body += a
				count -= 1
				if count != 0:
					message_body += ", "
			message_body += " - Need to be uploaded manually)"
	else:
		message_body +="\n\t\tPath to LPS USAA failed folder is invalid! Please correct!"
		print("Path to LPS USAA failed folder is invalid! Please correct!")

#Lending Space - Failed Folders
def lending_space():
	global message_body
	#path = "\\\\vrdmzsftp02\\lps$\\FileNET\\Failed"
	path=get_path("lending_space")
	if os.path.isdir(path):
		count = find_count(path)
		message_body += "\n\t\tLending Space : " + str(count) + " Failed Files"
	else:
		message_body +="\n\t\tPath to Lending Space failed folder is invalid! Please correct!"
		print("Path to Lending Space failed folder is invalid! Please correct!")

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
	
#function to send the mail with the message_body
def sendEmail(subject='',body='',attachment=''):
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
	subject += now
	mail.Subject = subject
	mail.Body = body
	if attachment != '':
		mail.Attachments.Add(attachment)
	try:
		mail.Send()
		print("Successfully sent email")
	except Exception:
		print("Error: unable to send email")
	
#Build the status	
def build_status():
	global message_body
	message_body = "Hi Team,\n\tPlease find below, the status report of the failed folders\n"
	blitz()
	lending_space()
	lps_nsm()
	lps_usaa()
	venture_nsm()
	venture_usaa()
	walz()
	message_body += "\n\nThanks & Regards,\nImaging Support"

#function that defines the path of the directories in which the failed folders should be to be searched
print("Processing...")
get_folder_paths()
build_status()
sendEmail("Failed Folder Monitoring ",message_body)
a=input("Press enter to quit the process")
if a != '':
	exit()
