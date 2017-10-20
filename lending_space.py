import os.path,shutil
import datetime,os
import fnmatch

now = str(datetime.date.today())
#path = "\\\\vrdmzsftp02\\lps$\\FileNET\\Failed\\"
cur_wrk_dir=os.getcwd()+"\\"
archive_path=os.getcwd()+"\\Archives\\"+str(now)+"\\"
filename='lending_space_files_list.txt'
file_path=cur_wrk_dir+filename
to_mail_list=[]
cc_mail_list=[]

#fetching the path for Lending Space
def get_path():
	global path
	file=open(cur_wrk_dir+"Utility\\paths.txt","r")
	for line in file:
		path=line.split("-")[1]
	#print(path)
	
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

#Send mail
def sendEmail():
	import win32com.client as win32
	set_mail_list()
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	mail_cc=";".join(cc_mail_list)
	mail.CC = mail_cc
	mail.Subject = "Lending Space Failed Files "+ now
	mail.Attachments.Add(file_path)
	mail.Body = "Hi Team,\n\tPFA the list of failed files present under the Lending Space Failed folder. Please do the needful.\n\nThanks & Regards,\nImaging Support"
	try:
		mail.Send()
		print("\nSuccessfully sent email")
	except Exception:
		print("\nError: unable to send email")
		
if not os.path.exists(archive_path):
	os.makedirs(archive_path)

print("---------------------")
print("Lending Space Process")
print("---------------------")
print("\nProcessing...")
get_path()
	
#Create a file with list of failed files
lending_space_files = open(filename, 'w')
for f in os.listdir(path):
	if fnmatch.fnmatch(f, '*.zip'):
		lending_space_files.write("%s\n" % str(f))
		src=path+str(f)
		shutil.move(src,archive_path)

lending_space_files.close()
if os.stat(file_path).st_size == 0:
	print("\nNo files are present under Lending Space")
else:
	sendEmail()
	
shutil.move(file_path,archive_path)

a=input("\nPress enter to quit the process")
if a != '':
	exit()



