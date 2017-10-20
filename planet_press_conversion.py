import os
import shutil
import fnmatch
import datetime 
import openpyxl
import fileinput
import sys
import time

#Getting the mailIds from the file
def get_mail_list(file_name):
	global mail_list
	file=open(file_name,"r")
	for line in file:
		if ".com" in line:	
			mail_list.append(line)
	
#Setting up the mailing list
def set_mail_list():
	mail_path=os.getcwd()+"\\Utility\\mail_address.txt"
	get_mail_list(mail_path)	
		
#Send the status mail
def sendMail(email_body):
	global brand
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	set_mail_list()
	mail_to=";".join(mail_list)
	mail.To = mail_to
	sub = "Planet Press Correction Report - "+str(now)
	mail.Subject=sub
	mail.Body = email_body
	try:
		mail.send()
	except Exception:
		print('')

#Step 8 - Checking if correction data is present for the file
def file_correction(file,input_data):
	for a in input_data:
		if file[0:len(file)-4] in a['Filename'][0:len(a['Filename'])-4]:
			replacement_data="|"+a['ProductId']+"||"
			with fileinput.FileInput(exceptions_dir+file, inplace=True) as file1:
				for line in file1:
					print(line.replace("|||",replacement_data), end='')
	
			del(replacement_data)
			replacement_data="|"+a['DocType']+"|"		
			with fileinput.FileInput(exceptions_dir+file, inplace=True) as file1:
				for line in file1:
					print(line.replace("||",replacement_data), end='')
			return True
	return False

#Step 7 - Processing the files
def process_file(file,input_data,exceptions_dir,destination_dir):
	global email_body
	if (file_correction(file,input_data)):
		print(file + " - Corrected!")
		email_body+="\n\t"+file+" - Corrected!"
		zip_file=file[0:len(file)-4]+".zip"
		if os.path.isfile(exceptions_dir+zip_file):
			shutil.move(exceptions_dir+zip_file,destination_dir+zip_file)
			shutil.move(exceptions_dir+file,destination_dir+file)
			pass
		else:
			print("Zip file not found")
			email_body+="Zip File not found"
			
	else:
		print(file + " - Not Corrected! Input record not found!")
		email_body+="\n\t"+file+" - Not Corrected! Input record not found!"

#Step 6
def initial_process(brand,exceptions_dir,destination_dir):
	print("Processing " + brand + " Files.......")
	files=[]
	input_file_name=''
	output_file_name=''
	global input_dir
	global email_body
	for f in os.listdir(exceptions_dir):
			if fnmatch.fnmatch(f, "*.txt"):
				files.append(f)
				
	if(len(files)) == 0:
		email_body+="\n\tNo files present are under " + brand + " Planet Press Statements and Expcetions folder"
		print("\nNo files in " + brand + " Planet Press Statements and Expcetions directory")
		a=input("Press enter to quit the process")
		if a != '':
			sys.exit()
	else:
		if brand.lower() == "cooper":
			input_file_name = input_dir+"cooper.xlsx"
			output_file_name = input_dir+"Completed\\Cooper\\"+"cooper__"+now+".xlsx"
		elif brand.lower() == "usaa":
			input_file_name = input_dir+"usaa.xlsx"
			output_file_name = input_dir+"Completed\\USAA\\"+"usaa__"+now+".xlsx"
		else:
			pass
		
		input_data=[]
		#Fetching the corections from input folder
		if os.path.isfile(input_file_name):
			wb = openpyxl.load_workbook(input_file_name)
			sheet=wb.active
			for row in sheet:
				singleRecord={};
				singleRecord['Filename']=row[0].value
				singleRecord['ProductId']=row[1].value
				singleRecord['DocType']=row[2].value
				input_data.append(singleRecord)
			wb.close()	
			del(wb)
			
			for file in files:
				process_file(file,input_data,exceptions_dir,destination_dir)
				zip_file=exceptions_dir+file[0:len(file)-4]+".zip"
			
			check_files=[]
			for a in files:
				check_files.append(a[0:len(a)-4])
			
			for a in input_data:
				if a['Filename'][0:len(a['Filename'])-4] not in check_files:
					print(a['Filename'][0:len(a['Filename'])-4] + ".txt  - Not found in exceptions folder. But present in input record.")
					email_body+="\n\t"+a['Filename'][0:len(a['Filename'])-4] + ".txt - Not found in exceptions folder. But present in input record."
					
				
			shutil.copy(input_file_name,output_file_name)
			os.remove(input_file_name)
			
			email_body+="\n\tThe Planet Press correction process has been completed for "+ brand +"."
			
		else:	
			print("\nInput File Not Found for " + brand + "!\nNote that the input file name should be simply \""+brand.lower()+".xlsx\"\n")
			email_body+="\nInput File Not Found for " + brand + "!\nNote that the input file name should be simply \""+brand.lower()+".xlsx\""
						
#Step 2 & 3 Read Paths
def read_paths(paths_file_location):
	global paths
	paths_file=open(paths_file_location,"r")
	for line in paths_file:
		paths.append(line.split(" - ")[1])
	paths_file.close()
	del(paths_file_location)
	del(paths_file)
	
#Step 1 - Declarations
paths=[]
brands=[]
exceptions_dirs=[]
destination_dirs=[]
email_body=''
input_dir=os.getcwd()+"\\Input\\"
now = str(datetime.datetime.now().strftime("%y_%m_%d__%H_%M"))
mail_list=[]

#Step 2 - Read the paths for Cooper
cooper_paths_file_location=os.getcwd()+"\\Utility\\cooper_paths.txt"
read_paths(cooper_paths_file_location)

#Step 3 - Read the paths for Cooper
usaa_paths_file_location=os.getcwd()+"\\Utility\\usaa_paths.txt"
read_paths(usaa_paths_file_location)

#Step 4 - Store the path and brand details in a single variable
brands.append(paths[0][0:len(paths[0])-1])
exceptions_dirs.append(paths[1][0:len(paths[1])-1])
destination_dirs.append(paths[2])
brands.append(paths[3][0:len(paths[3])-1])
exceptions_dirs.append(paths[4][0:len(paths[4])-1])
destination_dirs.append(paths[5])

email_body+="Hi Team,"
#Setp 5 - Process the domains one by one
count=0
for a in brands:
	print('')
	brand=a
	exceptions_dir=exceptions_dirs[count]
	destination_dir=destination_dirs[count]
	count+=1
	email_body+="\n\n"+brand
	initial_process(brand,exceptions_dir,destination_dir)
	
email_body+="\n\nThanks & Regards,\nImaging Support"
sendMail(email_body)
print("\nThe correction process is completed")
a=input("Press enter to quit the process")
if a != '':
	exit()
			
	
	
	

