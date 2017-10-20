import pyodbc
import os

ticket_number=''
type=''
type_tag=False
cur_wrk_dir=os.getcwd()
utility_dir=''

#Fetch Ticket Number
def fetch_ticket_number():
	global ticket_number
	ticket_number=input("Ticket Number - ")

#Fetch Ingestion Type
def fetch_ingestion_type():
	global type
	global type_tag
	global cur_wrk_dir
	global utility_dir
	type_option=input("Choose ingestion type...\n1)DMS\n2)Summit\nEnter choice - ")
	if type_option == "1":
		type="DMS"
		utility_dir=cur_wrk_dir+"\\Utility\\DMS\\"
		type_tag=True
	elif type_option == "2":
		type="Summit"
		utility_dir=cur_wrk_dir+"\\Utility\\Summit\\"
		type_tag=True
	else:
		print("Invalid type")
		type_tag=False

#Read File
def read_file(filename):
	data=''
	global ticket_number
	with open(filename, 'r') as myfile:
		data=myfile.read().replace('\n', '')
	query=data.replace('replace_ticket_number_here',ticket_number).strip()
	return query
			
def create_table(type):
	filename=utility_dir+"1_create_table.txt"
	query=read_file(filename)
	print(query)
	print("=================================================")
	
def bulk_insert(type):
	filename=utility_dir+"2_bulk_insert.txt"
	query=read_file(filename)
	print(query)
	print("=================================================")
	
def verify(type):
	filename=utility_dir+"3_verify_table.txt"
	query=read_file(filename)
	print(query)
	print("=================================================")
	
def insert_into_doc_admin(type):
	filename=utility_dir+"4_insert_into_doc_admin.txt"
	query=read_file(filename)
	print(query)
	print("=================================================")
	
def verify_doc_admin(type):
	filename=utility_dir+"5_verify_doc_admin.txt"
	query=read_file(filename)
	queries=query.split(";")
	print("Query 1 - " ,queries[0])
	print("Query 2 - " ,queries[1])
	print("=================================================")

#Step 1 - Fetch Ticket Number
fetch_ticket_number()
#Step 2 _ Fetch the type of ingestion
while(type_tag != True):
	fetch_ingestion_type()

print("--------------------------------------------------")
print("Ticket Number ------ ", ticket_number)
print("Ingestion Type ----- ", type)
print("--------------------------------------------------")

confirmation_tag=input("Can we proceed with the given input?\n1)Yes\n2)No\nYour choice...")
if confirmation_tag == "1":
	print("Processing.......")
	create_table(type)
	bulk_insert(type)
	verify(type)
	insert_into_doc_admin(type)
	verify_doc_admin(type)
	print("Process Completed")
else:
	print("Qutting the process")
