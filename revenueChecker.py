import openpyxl
from datetime import datetime
import shutil
import os, stat
from send_email import send_mail


def subAcctIdValidation(content, platform):

	content = str(content)
	stringSplitted =  content.split("-")

	if len(stringSplitted) != 2 or len(stringSplitted[0]) != 3:
		return False
	elif stringSplitted[0] in platform.keys():
		if platform[stringSplitted[0]] == 'accepted':
			return True
		else:
			return 'Warning'
	else:
		return False

def dateFormatValidation(content):
	content = str(content)
	try:
		dateTime = datetime.strptime(content, '%Y%m%d')
		if datetime.strptime('20100101', '%Y%m%d') <= dateTime <= datetime.strptime('20200101', '%Y%m%d'):
			return True
		return False

	except:
		return False


def productName(folderPath,fileName):

	book = openpyxl.load_workbook(folderPath+fileName)

	try:
		wb = book['Dim_List']

		rowIndex = 0
		colIndex = 0

		productList = []

		rows = wb.rows
		for row in rows:
			colIndex = 0
			if rowIndex == 0:
				rowIndex += 1
				continue
			for cell in row:
				if colIndex == 1:
					productList.append(cell.value)
				colIndex += 1

		return productList
	
	except:
		return []


def productNameValidation(content, productionNameList):
	return content in productionNameList

def searchTransTypeValidation(content):

	return "-" in str(content)

def revenueUnitFormatValidation(content):

	try:
		float(content)

	except:
		return False

	return True

def readPlatform(folderPath):

	platform ={}
	platformFile = open(folderPath+'platforms.txt','r')
	line = platformFile.readline()
	line = platformFile.readline().strip(' \t\n\r')

	while line:
		line = line.split("\t")
		platform[line[0].strip(' \t\n\r')] = line[1].strip(' \t\n\r')
		line = platformFile.readline().strip(' \t\n\r')

	platformFile.close()
	return platform



def main():

	targetFolderPath = 'D:/red/data/inbound/manual_adjustment/'

	fileName = "Essbase_Mth_End_Revenue_Adjustment_template.xlsx"

	todayDate = datetime.now().strftime("%Y%m%d")

	folderPath = targetFolderPath + todayDate + "/"

	if os.path.isdir(folderPath) == False:
		print("Folder in today's date " + todayDate + " does not exist! Program Terminated.")
		return

	
	
	logFile = open(folderPath+'log.txt','w')
	
	errorFile = open(folderPath+'Error_revenue_adjustment_'+todayDate+'.txt','w')
	
	# outputFile = open(folderPath+fileName.split('.')[0]+'.txt','w')
	outputFile = open(folderPath+'revenue_adjustment_'+todayDate+'.txt','w')

	platform = readPlatform(targetFolderPath)


	book = openpyxl.load_workbook(folderPath+fileName, data_only = True)

	if "Input" not in book.get_sheet_names():
		print('Input Tab Is Not Found in File! Program Terminated!')
		logFile.write('Input Tab Is Not Found in File! Program Terminated!'+'\n')
		logFile.close()
		return
	else:
		wb = book['Input']



	header = []
	rowIndex =  colIndex = 0
	noMissingCol = [0,1,2,3,5,6,9,13]
	productionNameList = productName(folderPath, fileName)
	subAcctIdWarning = []
	subAcctIdError = []


	rows = wb.rows

	for row in rows:

		if rowIndex == 0:
			logFile.write("Reading Header." + "\n")	

		colIndex = 0
		err = []
		tableCells = ''
		logFile.write("Line " + str(rowIndex+1) + ": " +"\n")
		
		tableContent = []
			
		for cell in row:
			tableContent.append(cell.value)
		
		if all(map(lambda x: x == tableContent[0], tableContent)):
			print('This line is probably empty. Skipped!')
			logFile.write('This line is probably empty. Skipped!' + '\n')
			rowIndex +=1
			continue				
				
		
		for col in tableContent:

			if rowIndex == 0:

				header.append(col)

			else:
				if colIndex in noMissingCol and  col is None:
					err.append('Error: ' + header[colIndex]+ " is empty."+"\n")
				

				if colIndex == 3 and not productNameValidation(col, productionNameList):
					logFile.write("Warning: " + "Production Name " + str(col) + "is not in provided list" + "\n")
					

				if colIndex == 0:
					if not subAcctIdValidation(col, platform):
						# logFile.write("Line " + str(rowIndex+1) + ": " + header[colIndex] + " has error. Skipped..." + "\n")
						err.append('Error: ' + header[colIndex] + ': ' + str(col)  + " has error." + "\n")
						subAcctIdError.append(str(col))
					elif subAcctIdValidation(col, platform) == "Warning":
						logFile.write("Warning: " + "Platform ID " + str(col) + "is valid but in exception" + "\n")
						subAcctIdWarning.append(str(col))


				if colIndex == 2 and not dateFormatValidation(col):
					err.append('Error: ' + header[colIndex] + ': ' + str(col) + " is not in correct date format." + "\n")

				if colIndex in [4,5] and searchTransTypeValidation(col):
					# logFile.write("Line " + str(rowIndex+1) + ": " + header[colIndex] + " has error. Skipped..." + "\n")
					err.append('Error: ' + header[colIndex] + ': ' + str(col) + " is not in correct search trans type." + "\n")

				if colIndex == [8,9] and not revenueUnitFormatValidation(col):
					# logFile.write("Line " + str(rowIndex+1) + ": " + header[colIndex] + " has error. Skipped..." + "\n")
					err.append('Error: ' + header[colIndex] + ': ' + str(col) + " is not in correct revenue unit format." + "\n")

			
			if col == None:
				tableCells += '' + "\t"
			else:
				tableCells += str(col) + "\t"

			colIndex += 1

		if len(header) != 14:
			print('The file contains more than 14 columns. Please check!')
			logFile.write('The file contains more than 14 columns. Please check!'+'\n')
			logFile.close()
			return
			
		if len(err) == 0:
			logFile.write("Okay!" + "\n")
			outputFile.write(tableCells + "\n")

		else:
			if len(subAcctIdError) > 0:
				outputFile.write(tableCells + "\n") # if platform has error, output to both error file and output file.
			errorFile.write('\t'.join(header) + '\n') #add header to error file
			errorFile.write(tableCells + '\n')			
			print(err)
			for e in err:
				logFile.write(e)


		rowIndex +=1

	logFile.close()
	errorFile.close()
	outputFile.close()
	book.close()
	
	#move text file folder to landing zone folder
	os.chmod(folderPath, 0o777)
	targetPath = 'D:/red/data/inbound/manual/adhoc/working/'+todayDate
	if not os.path.exists(targetPath):
		os.makedirs(targetPath)
	else:
		os.chmod(targetPath, 0o777)

	try:
		absolutePath = folderPath + 'revenue_adjustment_'+todayDate+'.txt'
		shutil.copy2(absolutePath, targetPath)
	except:
		print("Error in copy txt file to landing zone folder!")




	#send email to nofity user
	sentFrom = 'yang.xu@lexisnexisrisk.com'
	sentTo = ['Karen.Norero@lexisnexisrisk.com']
	sendCC = ['yang.xu@lexisnexisrisk.com']
	#If the platform is part of the optional list, send an email with these information:
	if len(subAcctIdWarning) > 0:
		subject = 'Revenue Adjustment Process Warning – Platform exception included'
		content = 'There are accounts in the file that have are part of the exception list.\n'
		for i in subAcctIdWarning:
			content = content + i + "\n"
		content += 'These accounts are included on the file to be processed on HPCC. Please validate before loading the data.'
		send_mail(sentFrom,sentTo,sendCC, subject,content)
	
	if len(subAcctIdError) > 0:
		subject = 'Revenue Adjustment Process Error – Invalid Platform included'
		content = 'There are accounts in the file that are part of the not to process list.\n'
		for i in subAcctIdError:
			content = content + i + "\n"
			print(content)
		content += 'These accounts are NOT included on the file to be processed on HPCC. Please validate the errors and reprocess these accounts.'
		print(content)
		send_mail(sentFrom,sentTo,sendCC, subject,content)
	if len(err) > 0:
		subject = 'Revenue Adjustment Process Error – General Error'
		content ='Errors have been found:' + "\n"
		for i in err:
			content = content + i
		send_mail(sentFrom,sentTo,sendCC, subject,content)


if __name__ =="__main__":
	main()





