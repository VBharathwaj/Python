import os.path,shutil
import datetime,os
import win32com.client
import fnmatch
import fileinput
from configparser import SafeConfigParser

now = str(datetime.date.today())
cur_wrk_dir=os.getcwd()+"\\"
input_folder_txt_files=[]
failed_folder_txt_files=[]
cooper_missing_txt_files=[]
usaa_missing_txt_files=[]
cooper_corrupted_files=[]
usaa_corrupted_files=[]
archive=os.getcwd()+"\\Archives\\"+str(now)
to_mail_list=[]
cc_mail_list=[]


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
def sendEmail(attachments,message):
	import win32com.client as win32
	set_mail_list()
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	mail_cc=";".join(cc_mail_list)
	mail.CC = mail_cc
	mail.Subject = "Venture Failed Files "+ now
	for attachment in attachments:
		mail.Attachments.Add(attachment)
	mail.Body = message
	try:
		mail.Send()
		print("\nSuccessfully sent email")
	except Exception:
		print("\nError: unable to send email")


#Writing the file list to the text file
def write_to_file(filepath,list):
	file=open(filepath,"w+")
	for a in list:
		file.write(a)
		file.write("\n")
	

#Get the count of files from the directory
def find_count(path):
    count = len([f for f in os.listdir(path)if os.path.isfile(os.path.join(path, f))])
    return count

#fetching the failed folder path for Lending Space
def get_paths(type):
	paths=[]
	parser = SafeConfigParser()
	parser.read(cur_wrk_dir+"Utility\\paths.ini")
	paths.append(parser.get(type, 'venture_nsm'))
	paths.append(parser.get(type, 'venture_usaa'))
	return paths

failed_folder_paths = get_paths('failed_paths')
input_folder_paths=get_paths('input_paths')
cooper_input_folder_path=input_folder_paths[0]
usaa_input_folder_path=input_folder_paths[1]

#Get the text files from input folder
def get_text_files_from_folder(path,type):
	global input_folder_txt_files
	global failed_folder_txt_files
	for f in os.listdir(path):
		if fnmatch.fnmatch(f, '*.txt'):
			if type == "input":
				input_folder_txt_files.append(str(f))
			elif type == "failed":
				failed_folder_txt_files.append(str(f))

#Check if the corresponding zip file exists and process
def process_file(file,failed_folder_path,input_folder_path,archive_path,brand):
	global input_folder_txt_files
	global failed_folder_txt_files
	global cooper_missing_txt_files
	global cooper_corrupted_files
	global usaa_missing_txt_files
	global usaa_corrupted_files
	txt_files=[]
	fname=file[0:len(file)-4]
	txt_file_name=fname+".txt"
	if txt_file_name in input_folder_txt_files:
		shutil.move(failed_folder_path+file,input_folder_path)
		pass
	elif txt_file_name in failed_folder_txt_files:
		if brand == "cooper":
			cooper_corrupted_files.append(file)
		elif brand == "usaa":
			usaa_corrupted_files.append(file)
		shutil.move(failed_folder_path+file,archive_path)
		shutil.move(failed_folder_path+txt_file_name,archive_path)
	else:
		if brand == "cooper":
			cooper_missing_txt_files.append(file)
		elif brand == "usaa":
			usaa_missing_txt_files.append(file)
	
print("------------------")
print(" Venture Process")
print("------------------")
		
for failed_folder_path in failed_folder_paths:
	failed_files=[]
	if "USAA" in failed_folder_path:
		print("\nProcessing USAA files")
		input_folder_path=usaa_input_folder_path
		archive_path=archive+"\\USAA\\"
		brand="usaa"
	else:
		print("\nProcessing Cooper files")
		input_folder_path=cooper_input_folder_path
		archive_path=archive+"\\Cooper\\"
		brand="cooper"
	
	count=find_count(failed_folder_path)
	if count == 0:
		print("\nNo failed files in Venture process")
		continue

	for f in os.listdir(failed_folder_path):
		if fnmatch.fnmatch(f, '*.zip'):
			failed_files.append(str(f))
		
	get_text_files_from_folder(input_folder_path,"input")
	get_text_files_from_folder(failed_folder_path,"failed")
	
	if not os.path.exists(archive_path):
		os.makedirs(archive_path)
		
	for single_file in failed_files:
		process_file(single_file,failed_folder_path,input_folder_path,archive_path,brand)
	
cooper_corrupted_file=archive+"\\Cooper\\cooper_corrupted_files.txt"
write_to_file(cooper_corrupted_file,cooper_corrupted_files)
usaa_corrupted_file=archive+"\\USAA\\usaa_corrupted_files.txt"
write_to_file(usaa_corrupted_file,usaa_corrupted_files)
cooper_missing_file=archive+"\\Cooper\\cooper_missing_files.txt"
write_to_file(cooper_missing_file,cooper_missing_txt_files)
usaa_missing_file=archive+"\\USAA\\usaa_missing_files.txt"
write_to_file(usaa_missing_file,usaa_missing_txt_files)

attachments=[]
if len(cooper_corrupted_files) > 0:
	attachments.append(cooper_corrupted_file)
if len(usaa_corrupted_files) > 0:
	attachments.append(usaa_corrupted_file)
if len(cooper_corrupted_files) > 0 or len(usaa_corrupted_files) > 0:
	message="Hi Team,\n\tSome files were corrupted under venture process. PFA, the list of file(s) failed during the Venture Process. Kindly retransmit the list of files mentioned in the attachment.\n\nThanks & Regards,\nImaging Support"
	sendEmail(attachments,message)
	print("\nCorrupted files were mailed for resubmission")
	
del(attachments)
attachments=[]
if len(cooper_missing_txt_files) > 0:
	attachments.append(cooper_missing_file)
if len(usaa_missing_txt_files) > 0:
	attachments.append(usaa_missing_file)
if len(cooper_missing_txt_files) > 0 or len(usaa_missing_txt_files) > 0:
	message="Hi Team,\n\tSome Text files were missing under venture process. PFA, the list of zip file(s) for which the text files are not found. Kindly retransmit the list of text files for the corresponding zip files mentioned in the attachment.\n\nThanks & Regards,\nImaging Support"
	sendEmail(attachments,message)
	print("\nMissing files were mailed for resubmission")
	
print("\nProcess completed")
		
a=input("\nPress enter to quit the process")
if a != '':
	exit()
