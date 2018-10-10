# Imports Required for the Application

import os
import csv
import pytz
import getpass
import logging
import psycopg2
import datetime
from pytz import timezone
from os.path import basename
from configparser import ConfigParser

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
	
# Function to Get Current Backup File Path

def getCurrentBackupFilePath(currentBackupFolderPath, currentDate):
	currentBackupFilePath = currentBackupFolderPath + "\\" + "Backup_" + currentDate.replace("-", "_") + "_v0.csv"
	if(os.path.isfile(currentBackupFilePath)):
		file_number = 0
		while True:
			file_number = file_number + 1
			temp_file_name = currentBackupFolderPath + "\\" + "Backup_" + currentDate.replace("-", "_") + "_v" + str(file_number) +".csv"
			if not os.path.isfile(temp_file_name):
				break
		return temp_file_name
	else:
		return currentBackupFilePath
	
# Function to Backup Current Data from the destination table 

def backupCurrentData(currentBackupFolderPath, tableName, cur, tableProps):
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
		
	currentBackupFilePath = getCurrentBackupFilePath(currentBackupFolderPath, currentDate)
	
	currentBackupFile = open(currentBackupFilePath, 'w', newline = '')  
	with currentBackupFile:  
	   writer = csv.writer(currentBackupFile)
	   writer.writerows(backupData)
	
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
logging.info("Inititing Backup Process")

# Application Declarations

databaseNames = []
backupFolderPath = currentWorkingDirectory + "\\Backup"

#Create Backup and Report folders if they does not exist

if not os.path.exists(backupFolderPath):
	os.makedirs(backupFolderPath)
	
# Parsing the Configuration File

parser.read(currentWorkingDirectory+"\\Configuration.ini")
destinationEnvironment = parser.get('backupProperties','destinationEnvironment') 
specificFlag = parser.get('backupProperties','specific')
migrationType = ("Single" if specificFlag.lower() == "true" else "Multiple")

# If Specific Flag is True, Database Details from Config File is fetched

if migrationType == "Single":
	databaseNames = parser.get('database','db_name').split(",\n")

# Parsing the Properties File - Fetching Database Usernames and Passwords:

parser.read(currentWorkingDirectory+"\\Config\\Properties.ini")
destinationServer = parser.get(destinationEnvironment,'hostname')
destinationDatabaseUserName = parser.get(destinationEnvironment,'db_username')
destinationDatabasePassword = parser.get(destinationEnvironment,'db_password')

# If Specific Flag is False, Database Details from Properties File is fetched

if migrationType == "Multiple":
	databaseNames = parser.get('database','db_name').split(",\n")

logging.info("Backup Type: " + migrationType)
	
# Starting the Process

for databaseName in databaseNames:
	tableNames = []
	tableNames = parser.get(('table' if migrationType == "Single" else databaseName),'table_name').split(",\n")
	destinationDatabaseName = databaseName + destinationEnvironment
	logging.info("Started backing up tables in " + destinationDatabaseName + " database")
	tableCount = 0
	
	# Database Connection
	
	conn = psycopg2.connect(host=destinationServer,database=destinationDatabaseName, user=destinationDatabaseUserName, password=destinationDatabasePassword)
	cur = conn.cursor()
	
	# Processing each Tables
	
	for tableName  in tableNames:
		tableCount = tableCount + 1
		
		tableProps = []
		logging.info("Started backing up " + tableName + " table")

		# Fetching Table Props
		
		columnNamesAndDataType = fetchTableProps(cur, tableName)
		columnNamesString = columnNamesAndDataType[0]
		columnNamesWithDataTypeString = columnNamesAndDataType[1]

		# Backup Current Data from Destination Database Table
		
		currentBackupFolderPath = backupFolderPath + "\\" + destinationDatabaseName + "\\" + currentDate + "\\" + tableName
		
		if not os.path.exists(currentBackupFolderPath):
			os.makedirs(currentBackupFolderPath)
			
		backupCurrentData(currentBackupFolderPath, tableName, cur, tableProps)
		logging.info("Completed backing up " + tableName + " table")
	logging.info("------------------------- DATABASE REPORT -------------------------")
	logging.info("Completed backing up table in " + destinationDatabaseName + " database")
	logging.info("Number of Tables Backed up in " + destinationDatabaseName + " database : " + str(tableCount))
	tableCount = 0
	cur.close()
	del(cur)
	conn.close()
	del(conn)
	logging.info("------------------------- END OF DATABASE -------------------------")
	
logging.info("Backup Process Complete")
logging.info("****************************************************END OF PROCESSING****************************************************")





























	
	