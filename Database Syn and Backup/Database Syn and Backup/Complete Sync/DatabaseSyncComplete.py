# Imports Required for the Application

import os
import csv
import pytz
import getpass
import smtplib
import logging
import psycopg2
import datetime
from pytz import timezone
from six import string_types
from bs4.builder import HTML
from os.path import basename
from email.mime.text import MIMEText
from configparser import ConfigParser
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate
from email.mime.application import MIMEApplication

# Function to Fetch the Column Names and Column Properties

def fetchTableProps(cur, 
				tableName):
				
	getTablePropsQuery = "SELECT column_name,data_type from information_schema.columns where table_name='" + tableName +"'"
	cur.execute(getTablePropsQuery)
	row = cur.fetchone()
	data = []
	columnNamesString = ''
	columnNamesWithDataTypeString = ''
	
	while row is not None:
		temp = {}
		temp['ColumnName'] = row[0]
		if "character" in row[1].lower(): 
			temp['DataType'] = "text" 
		else:
			temp['DataType'] = row[1]
		columnNamesString += "\"" + temp['ColumnName'] + "\""
		columnNamesWithDataTypeString += temp['ColumnName'] + " " + temp['DataType']
		tableProps.append(temp)
		del(temp)
		row = cur.fetchone()
		if row is not None:
			columnNamesString += ","
			columnNamesWithDataTypeString += ","
			
	data.append(columnNamesString)
	data.append(columnNamesWithDataTypeString)
	return data;
	
#Function to Form/Build the Migrate Query
	
def fromMigrateQuery(tableName, 
				columnNamesString, 
				sourceServer, 
				sourceDatabaseUserName, 
				sourceDatabasePassword, 
				sourceDatabaseName, 
				columnNamesWithDataTypeString):
				
	migrateQuery = "INSERT INTO \"" + tableName + "\"(" + columnNamesString + ") "
	migrateQuery += "SELECT " + columnNamesString.replace("\"", "") + " FROM "
	migrateQuery += "dblink('host=" + sourceServer + " user=" + sourceDatabaseUserName + " password=" + sourceDatabasePassword + " dbname=" + sourceDatabaseName + "',"
	migrateQuery += "'SELECT " + columnNamesString + " FROM public.\"" + tableName + "\"')"
	migrateQuery += " As " + (tableName[1:] if tableName[0].isdigit() else tableName)
	migrateQuery += "_v(" + columnNamesWithDataTypeString + ")"
	return migrateQuery

# Function to Get Current Backup File Path

def getFilePath(folderPath, fileType, currentDate):
	filePath = folderPath + "\\" + fileType + currentDate.replace("-", "_") + "_v0.csv"
	if(os.path.isfile(filePath)):
		file_number = 0
		while True:
			file_number = file_number + 1
			temp_file_name = folderPath + "\\" + fileType + currentDate.replace("-", "_") + "_v" + str(file_number) + ".csv"
			if not os.path.isfile(temp_file_name):
				break
		tempData = []
		tempData.append(temp_file_name)
		tempData.append(str(file_number))
		return tempData
	else:
		tempData = []
		tempData.append(filePath)
		tempData.append(str(0))
		return tempData
	
# Function to Backup Current Data from the destination table 

def backupCurrentData(destinationEnvironment, currentBackupFolderPath, tableName, cur, tableProps):
	getDestinationTabledataQuery = "SELECT * from public.\"" + tableName + "\""	

	backupData = []
	temp = []
	
	for singleProp in tableProps:
		temp.append(singleProp['ColumnName'])
	backupData.append(temp)
	
	cur.execute(getDestinationTabledataQuery)
	row = cur.fetchone()
	
	while row is not None:
		temp = []
		for singleValue in row:
			if tableName == "lsphere_lookup_salesitem":
				temp.append(singleValue.encode("utf-8") if isinstance(singleValue, str) else singleValue)
			else:
				temp.append(singleValue)
		backupData.append(temp)
		row = cur.fetchone()
		
	currentBackupFilePath = getFilePath(currentBackupFolderPath, destinationEnvironment + "_Backup_", currentDate)[0]
	
	currentBackupFile = open(currentBackupFilePath, 'w', newline = '')  
	with currentBackupFile:  
	   writer = csv.writer(currentBackupFile)
	   writer.writerows(backupData)
	
	return currentBackupFilePath
	
# Send SMTP Mail	

def sendSMTPEmail(resultMail, attachment):
	global to_mail_list
	global cc_mail_list
	global smtpHostname
	global smtpPortNumber
	# global smtpUsername
	# global smtpPassword
	
	mail_to=";".join(to_mail_list)
	mail_cc=";".join(cc_mail_list)
	sub = sourceEnvironment + " to " + destinationEnvironment + " Data Migration Report"
	msg = MIMEMultipart('alternative')
	msg['Subject'] = sub
	msg['From'] = 'WIA_Data_Migrations@xome.com'
	msg['To'] = mail_to
	msg['cc'] = mail_cc
	part = MIMEText(resultMail, 'html')
	msg.attach(part)
	
	with open(attachment, "rb") as fil:
		part2 = MIMEApplication(
			fil.read(),
			Name=basename(attachment)
		)
	part2['Content-Disposition'] = 'attachment; filename="%s"' % basename(attachment)
	msg.attach(part2)
	
	s = smtplib.SMTP(smtpHostname, smtpPortNumber)
	#s.login(smtpUsername, smtpPassword)
	s.sendmail('WIA_Data_Migrations@xome.com', 'Bharathwaj.Vasudevan@xome.com', msg.as_string())
	s.quit()
	
# Global Declarations

currentWorkingDirectory=os.getcwd()
currentUser = getpass.getuser()
currentIndianTimeZone = timezone('Asia/Calcutta')
currentDate = (str(datetime.date.today()).replace(".","_"))
currentTime = str(datetime.datetime.time(datetime.datetime.now()))[:8]
currentIndianDateTime = str((datetime.datetime.now()).astimezone(currentIndianTimeZone))
currentIndianDate = currentIndianDateTime[:10]
currentIndianTime = currentIndianDateTime[11:19]
parser = ConfigParser()

# Initiating Logging

logsFolderPath = currentWorkingDirectory + "\\Logs"
if not os.path.exists(logsFolderPath):
	os.makedirs(logsFolderPath)
logging.basicConfig(filename=logsFolderPath + '\\Logs.log',level=logging.DEBUG, format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p')
logging.info("***************************************************START OF PROCESSING***************************************************")
logging.info("Initialising application")

# Application Declarations

databaseNames = []
backupFolderPath = currentWorkingDirectory + "\\Backup"
reportFolderPath = currentWorkingDirectory + "\\Report"
reportProperty = ''
reportData = []
reportHeader = ['Destination Database',
				'Source Database',
				'Table Name',
				'Records Deleted',
				'Records Inserted',
				'Backup Location']

#Create Backup and Report folders if they does not exist

if not os.path.exists(backupFolderPath):
	os.makedirs(backupFolderPath)
	logging.info("Backup Folder does not exist. Creating backup folder")
	
if not os.path.exists(reportFolderPath):
	os.makedirs(reportFolderPath)
	logging.info("Report Folder does not exist. Creating report folder")


migrationType = "Multiple"
logging.info("Parsing Proprties file")
parser.read(currentWorkingDirectory+"\\Config\\Properties.ini")
environments = parser.get('environments','env').split(",\n")

for environment in environments:
	reportData = []
	sourceEnvironment = "Prod"
	destinationEnvironment = environment
	print('destinationEnvironment : '  + destinationEnvironment)

	if destinationEnvironment.lower() == "prod":
		print("Migration to Production database is not authorised")
		logging.info("Attempted migration to Production. Blocked the process")
	else:

		# Parsing the Properties File - Fetching Database Usernames and Passwords:

		destinationServer = parser.get(destinationEnvironment,'hostname')
		destinationDatabaseUserName = parser.get(destinationEnvironment,'db_username')
		destinationDatabasePassword = parser.get(destinationEnvironment,'db_password')

		sourceServer = parser.get(sourceEnvironment,'hostname')
		sourceDatabaseUserName = parser.get(sourceEnvironment,'db_username')
		sourceDatabasePassword = parser.get(sourceEnvironment,'db_password')

		# Parsing the Properties File - Fetching SMTP Properties:

		smtpHostname = parser.get('smtpProperties','hostname')
		smtpPortNumber = parser.get('smtpProperties','port')
		smtpUsername = parser.get('smtpProperties','username')
		smtpPassword = parser.get('smtpProperties','password')

		to_mail_list = parser.get('mailID','to').split(';')
		cc_mail_list = parser.get('mailID','cc').split(';')

		# If Specific Flag is False, Database Details from Properties File is fetched

		databaseNames = parser.get('database','db_name').split(",\n")
		logging.info("Migration Type: Multiple  - Initiated")
			
		logging.info("Building data for report.txt")
		# Build Report Property Text File Data
		reportProperty = reportProperty + "Date(CST/IST): " + currentDate + " | " + currentIndianDate + "\n"
		reportProperty = reportProperty + "Time(CST/IST): " + currentTime + " | " + currentIndianTime + "\n"
		reportProperty = reportProperty + "User: " + currentUser + "\n\n"

		reportProperty = reportProperty + "Environment Details\n"
		reportProperty = reportProperty + "Source Environment: " + sourceEnvironment + "\n"
		reportProperty = reportProperty + "Destination Environment: " + destinationEnvironment + "\n"
		reportProperty = reportProperty + "Migration: " + migrationType  + "\n\n"

		reportProperty = reportProperty + "Destination Details\n"
		reportProperty = reportProperty + "Server: " + destinationServer + "\n"
		reportProperty = reportProperty + "Username: " + destinationDatabaseUserName + "\n"
		reportProperty = reportProperty + "Password: " + destinationDatabasePassword + "\n\n"

		reportProperty = reportProperty + "Source Details\n"
		reportProperty = reportProperty + "Server: " + sourceServer + "\n"
		reportProperty = reportProperty + "Username: " + sourceDatabaseUserName + "\n"
		reportProperty = reportProperty + "Password: " + sourceDatabasePassword + "\n"

		reportData.append(reportHeader)

		# Starting the Process

		logging.info("Starting the Migration")
		for databaseName in databaseNames:
			tableNames = []
			tableNames = parser.get(('table' if migrationType == "Single" else databaseName),'table_name').split(",\n")
			sourceDatabaseName = databaseName + sourceEnvironment
			destinationDatabaseName = databaseName + destinationEnvironment
			
			# Processing each Tables
			logging.info("Started processing the database " + destinationDatabaseName)
			for tableName  in tableNames:
				conn = psycopg2.connect(host=destinationServer,database=destinationDatabaseName, user=destinationDatabaseUserName, password=destinationDatabasePassword)
				cur = conn.cursor()
				logging.info("Started processing the table " + tableName + " in " + destinationDatabaseName + " database")

				tableProps = []
				reportDataTemp = []
				
				# Fetching Table Props
				
				logging.info("Fetching the column names and data types")
				columnNamesAndDataType = fetchTableProps(cur, tableName)
				columnNamesString = columnNamesAndDataType[0]
				columnNamesWithDataTypeString = columnNamesAndDataType[1]

				# Backup Current Data from Destination Database Table
				
				logging.info("Backing up current data from " + tableName + " in " + destinationDatabaseName + " database")
				currentBackupFolderPath = backupFolderPath + "\\" + currentDate + "\\" + destinationDatabaseName + "\\" + tableName
				
				if not os.path.exists(currentBackupFolderPath):
					os.makedirs(currentBackupFolderPath)
					
				currentBackupFilePath = backupCurrentData(destinationEnvironment, currentBackupFolderPath, tableName, cur, tableProps)
				logging.info("Backup process complete")
				
				# Delete Existing Data from Destination Database Table
				
				logging.info("Deleting current data from " + tableName + " in " + destinationDatabaseName + " database")
				deleteDestinationTableDataQuery = "DELETE FROM \"" + tableName + "\""
				cur.execute(deleteDestinationTableDataQuery)
				deleteCount  = str(cur.rowcount)
				logging.info("Deletion process complete")
				conn.commit()
				
				# Form Migrate Query
				
				logging.info("Started migrating data from " + tableName + " in " + sourceDatabaseName + " to " + destinationDatabaseName + " database")
				migrateQuery = fromMigrateQuery(tableName, columnNamesString, sourceServer, sourceDatabaseUserName, sourceDatabasePassword, sourceDatabaseName, columnNamesWithDataTypeString)
				
				# Executing the Data Migration Query
				
				cur.execute(migrateQuery)
				conn.commit()
				insertCount = str(cur.rowcount)
				logging.info("Migration of data complete")
				
				# Append Data to Report Data
			
				reportDataTemp.append(destinationDatabaseName)
				reportDataTemp.append(sourceDatabaseName)
				reportDataTemp.append(tableName)
				reportDataTemp.append(deleteCount)
				reportDataTemp.append(insertCount)
				reportDataTemp.append(currentBackupFilePath)
				reportData.append(reportDataTemp)
				deleteCount = 0
				insertCount = 0
				cur.close()
				conn.close()
		
		# Fetch Report File Path

		currentReportFolderPath = reportFolderPath + "\\" + currentDate + "\\" + destinationDatabaseName + "\\" + tableName
		if not os.path.exists(currentReportFolderPath):
			os.makedirs(currentReportFolderPath)
			
		logging.info("Writing migration details to report.csv file")
		reportFilePathData = getFilePath(currentReportFolderPath, destinationEnvironment + "_Report_" , currentDate)
		currentReportFilePath = reportFilePathData[0]

		currentReportFile = open(currentReportFilePath, 'w', newline = '')  
		with currentReportFile:  
		   writer = csv.writer(currentReportFile)
		   writer.writerows(reportData)	

		logging.info("Writing to report.csv complete")
		logging.info("Writing connection details  to report.txt file")
		currentReportPropertyFilePath= currentReportFolderPath + "\\" + destinationEnvironment + "_Report_" + currentDate.replace("-", "_") + "_v" +  str(reportFilePathData[1]) + ".txt"
		currentReportPropertyFile = open(currentReportPropertyFilePath, 'w')
		currentReportPropertyFile.write(reportProperty)
		currentReportPropertyFile.close()
		logging.info("Writing to report.txt file complete")

		# Form Report Mail Body

		logging.info("Building mail body")
		reportMailTable = ''
		reportMailTable += "<table border = "'1'">"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "Process Date (CST/IST)" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" +  currentDate + " | " + currentIndianDate + "</td>"
		reportMailTable+="</tr>"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "Process Time (CST/IST)" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" + currentTime + " | " + currentIndianTime + "</td>"
		reportMailTable+="</tr>"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "User Name" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" + currentUser + "</td>"
		reportMailTable+="</tr>"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "Source Environment" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" + sourceEnvironment + "</td>"
		reportMailTable+="</tr>"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "Destination Environment" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" + destinationEnvironment + "</td>"
		reportMailTable+="</tr>"
		reportMailTable+="<tr>"
		reportMailTable+="<td bgcolor="'#2196F3'" height="'50'" width="'300'" align="'center'">" + "Migration Type" + "</td>"
		reportMailTable += "<td height="'50'" width="'300'" align="'center'">" + migrationType + "</td>"
		reportMailTable+="</tr>"
		reportMailTable +="</table>"

		# Send Email

		reportMailBody='<html><body>Hi Team,<br><br>Data Migration process is complete. PFA, the Migration Report<br><br>'+reportMailTable+'<br>Thanks,<br>Bharath</body></html>'
		
		logging.info("Sending out the report mail")
		sendSMTPEmail(reportMailBody, currentReportFilePath)
logging.info("Entire process completed successfully!")
logging.info("****************************************************END OF PROCESSING****************************************************")