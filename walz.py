import os.path,shutil
import datetime,os
from shutil import move
import fnmatch
import time
import win32com.client
from zipfile import ZipFile, ZIP_DEFLATED
import fileinput
from contextlib import closing

now = str(datetime.date.today())
#now="2017-09-11"
cur_wrk_dir=os.getcwd()+"\\"
corrupted_files=[]
zip_dir=''

#search_dir="\\\\vrisi01\\users$\\bvasudeva1\\Windows\\Desktop\\Walz\\";
search_dir="\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\WALZ\\Failed\\"

#duplicates_dstn="\\\\vrisi01\\users$\\bvasudeva1\\Windows\\Desktop\\Walz\\Duplicates\\"
duplicates_dstn="\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\WALZ\\Failed\\Duplicates"

walz_dir="\\\\vrisi01\\users$\\bvasudeva1\\Windows\\Desktop\\winpput\\"
#walz_dir="\\\\vrisi01\\shared$\\rshare\\checit\\imaging\WALZ\\"

#Archive Paths
corrupt_archive_path=cur_wrk_dir+"\\Archives\\Walz\\Corrupted\\"+str(now)+"\\"
pipe_delimiter_archive_path=cur_wrk_dir+"\\Archives\\Walz\\Pipe Delimited\\"+str(now)+"\\"
duplicates_archive_path=cur_wrk_dir+"\\Archives\\Walz\\Duplicates\\"+str(now)+"\\"

def file_exists_error(search_dir,filename):
	global dstn_input
	src=search_dir+filename
	dstn=duplicates_dstn+filename
	shutil.copy(src,duplicates_archive_path)
	move(src,dstn)

def pipeline_rezip_folder(basedir, archivename):
    assert os.path.isdir(basedir)
    with closing(ZipFile(archivename, "w", ZIP_DEFLATED)) as z:
        for root, dirs, files in os.walk(basedir):
            #NOTE: ignore empty directories
            for fn in files:
                absfn = os.path.join(root, fn)
                zfn = absfn[len(basedir)+len(os.sep):] #XXX: relative path
                z.write(absfn, zfn)
    shutil.move(zfn,archivename)
	
def pipeline_replace(text_to_be_replaced,replacement_text,text_file_path):
#Replace the text
	with fileinput.FileInput(text_file_path, inplace=True) as file:
		for line in file:
			print(line.replace(text_to_be_replaced,replacement_text), end='')
	
def pipeline_replace_preprocess(srcdir,filename,content):
	words_in_line=[];
	words_in_line=content.split("|")
	text_file_path=srcdir+filename[0:len(filename)-4]+".txt"
	replacement_text=str(words_in_line[2])[0:len(words_in_line[2])-1]+" "+str(words_in_line[3])[1:]
	if words_in_line[2] + "|" + words_in_line[3] in content:
		text_to_be_replaced=words_in_line[2] + "|" + words_in_line[3]
		pipeline_replace(text_to_be_replaced,replacement_text,text_file_path)
		
	elif words_in_line[2] + " | " + words_in_line[3] in content:
		text_to_be_replaced=words_in_line[2] + " | " + words_in_line[3]
		pipeline_replace(text_to_be_replaced,replacement_text,text_file_path)

def pipeline_delimiter_error(search_dir,filename,content):
	global zip_dir
	basedir=search_dir+filename[0:len(filename)-4]
	zip_ref = ZipFile(search_dir+filename, 'r')
	extract_path=cur_wrk_dir+filename[0:len(filename)-4]+"\\"
	if not os.path.exists(extract_path):
		os.makedirs(extract_path)
	zip_ref.extractall(extract_path)
	zip_ref.close()
	print("Unzipped")
	dstndir=pipe_delimiter_archive_path
	currentdir=cur_wrk_dir+filename[0:len(filename)-4]+".zip"
	shutil.copy(search_dir+filename,dstndir+filename)
	pipeline_replace_preprocess(extract_path,filename,content)
	print("Text Replaced")
	zip_dir=cur_wrk_dir+filename[0:len(filename)-4]
	shutil.make_archive(filename[0:len(filename)-4], 'zip', zip_dir)
	print("Zipped again!");
	os.remove(search_dir+filename)
	shutil.copy(currentdir,walz_dir)
	os.remove(currentdir)
	print("Moved for reproceesing")	
		
def corrupted_file_error(filename,search_dir):
	global corrupted_files
	corrupted_files.append(filename)
	move(search_dir+filename,corrupt_archive_path)
		
def write_to_file():
	the_file = open('corrupted_files.txt', 'w')
	for a in corrupted_files:
		the_file.write("%s\n" % a)
	the_file.close()
				
def sendEmail(subject='',body='',attachment=''):
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	#emails = ";".join(emailList)
	mail.To = "balaji.s@mrcooper.com"
	mail.CC = "bharathwaj.vasudevan@mrcooper.com"
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
	
walz_failed_files=[];
print("Processing...")
if not os.path.exists(corrupt_archive_path):
	os.makedirs(corrupt_archive_path)
if not os.path.exists(pipe_delimiter_archive_path):
	os.makedirs(pipe_delimiter_archive_path)
if not os.path.exists(duplicates_archive_path):
	os.makedirs(duplicates_archive_path)
	
for f in os.listdir(search_dir):
	if fnmatch.fnmatch(f, '*.zip'):
		single_file={};
		single_file['Filename']=f
		single_file['Reason']=''
		single_file['Content']=''
		walz_failed_files.append(single_file)
		del(single_file)
		
print("No. of failed files:",len(walz_failed_files))

#Search reason from outlook		
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
search_folder = "App Trace"
outlook = outlook.Folders

#Fetch messages from outlook
for primary in outlook:
	for secondary in primary.Folders:
		if str(secondary) == search_folder:
			messages=secondary.Items 
print("Fetched Mails")
					
#Fetch the reason for each file
for message in messages:
	message_date=str(message.ReceivedTime)[0:10]
	if now == message_date:
		content=str(message.body)[130:720]
		for b in walz_failed_files:
			file_name=str(b['Filename'])
			if file_name in content:
				if "pipe-delimited index" in content:
					b['Reason']+="Pipe Delimited Index"
					b['Content']=content[250:650]
					break
				elif "file exists" in content:
					b['Reason']+="File already exists"
					break
				elif "used by another process" in content or "timeout " in content:
					break
				else:
					b['Reason']+="Corrupted File"
					break			
print("Fetched Failed Reasons") 


print("Processing Files")
count=1
for single_file in walz_failed_files:
	print("Processed %d out of %d files",count,len(walz_failed_files))
	count+=1
	fname=str(single_file['Filename'])
	fname=fname[0:len(fname)-4]
	if "Corrupted File" in single_file['Reason']:
		corrupted_file_error(single_file['Filename'],search_dir)
		
	elif "Pipe Delimited Index" in single_file['Reason']:
		pipeline_delimiter_error(search_dir,single_file['Filename'],single_file['Content'])
		shutil.move(cur_wrk_dir+fname,pipe_delimiter_archive_path)

	elif "File already exists" in single_file['Reason']:
		file_exists_error(search_dir,single_file['Filename'])

		
if len(corrupted_files) > 0:
	write_to_file()
	attachment_location=cur_wrk_dir+"corrupted_files.txt"
	lending_space_body="PFA list of files which failed during the walz transfer process. These files are corrupted. Kindly redrop these files."
	sendEmail("Walz Corrupted Files ",lending_space_body,attachment_location)
	os.remove(attachment_location)
print("Done!")