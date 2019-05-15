# -*- coding: utf-8 -*-
"""
Created on Fri May 10 13:04:05 2019

The purpose of the file to complete the first step of paying property tax amounts.

More specifically the first step is to create a job file in the proper folder. 
NJ181818 would be created at the file within CTA Paid Taxes

@author: austin.schrader
"""

import os
import shutil
import os.path

# ======================== Read in a file name, create the file ==============================
# Ask for the job number that the payer would like to start processing for payment
jobRequested = input("What property tax job would you like to process? ")

# Add in the necessary backslashes to a) add a backslash, 2) provide the escape. Ex: \NJ484848
fileName = "\\" + jobRequested

# Read the jobRequested's stateCode and append that state to the path
stateCode = "\\" + fileName[1:3]

# Append the inputted fileName 
path = "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019\\" + stateCode + fileName

# Make the file path, and tell the user if there was some issue
try:
    os.mkdir(path)
except OSError:
    print("Creation of the directory %s failed" % path)
else:
    print("Succesfully created the directory %s " % path)
    
# ============================= Add in the supporting files ==================================    

print("Adding supporting files...")

# Grab the supporting files locations and save their paths as a variable
sourceBillExport = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\Bill export B Tool.xlsm"
sourceDoNotPay = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\For Russell to review-Loan Template.xlsm"
sourceForRussell = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\INTACCT Escrow Upload Template.xlsm"

# Pass in the supporting files locations and save them in the job that we're currently doing
shutil.copyfile(sourceBillExport, "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\Bill export B Tool.xlsm")
shutil.copyfile(sourceDoNotPay, "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\For Russell to review-Loan Template.xlsm")
shutil.copyfile(sourceForRussell, "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\INTACCT Escrow Upload Template.xlsm")

print("Support files added!")

# ========== Open Lereta, Download Job's Exceptions, main PDF, and main Excel =================

print("Add these files to the job file: main PDF from Lereta's job, Excel sheet, and PDF exceptions.")

input("Once those files are added, PRESS any key to continue...")
# Requests, Beautiful soup to open Lereta's pages with what the url should be
# Then, navigate to Exception's URL and download the PDF if there is one into the job's working folder
# Then navigate to the disbursement's URL and download the PDF if there is one into the job's working folder
# Then on that same page as the disbursement's URL, download the Excel file and save it into the job's working folder 

# ======================= Format the Export-TCS36501 Excel file ===============================

### COMPLETED - NEEDS INTEGRATION ### Open Export-TSC36501 - Completed
### COMPLETED - NEEDS INTEGRATION ### Copy Columns A, B, E, G, AF, AN into a new file called Export-Final
### COMPLETED - NEEDS INTEGRATION ### When choosing F - M, if O == "1" then answer = G. If O == "2" then answer = I, etc
### COMPLETED - NEEDS INTEGRATION ### Output an excel file that indicates the order of name of the counties, the installment year, and installment period

# To use the openpyxl module, we first have to convert the Excel file Export-TCS36501 from xls into xlsx file format
import win32com.client as win32
# Openpyxl is the module that's doing the editing of an excel. However, it can only wrok with .xlsx formats
#import openpyxl
import pandas

# Job fileName and job stateCode, hardcoded which should be replaced later
#fileName = "\\NJ666666"
#stateCode = "\\NJ"

# Opens the .xls file located at this location
fname = "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\Export-TCS36501.xls"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

# Saves the .xls file as .xlsx (ie, it converts the filetype to the filetype that we can use)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

# Establishes the path for the .xlsx file (the version that we can work with)
jobPath = ("\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName)
exportFilePath = "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\Export-TCS36501.xlsx"

# Reads the excel and parses out columns 0,1,4,6 etc (all the ones we need.)
# Then, it exports the file to the jobPath + \\output.xlsx
dataframe = pandas.read_excel(exportFilePath, parse_cols = [0, 1, 4, 6, 8, 10, 12, 14, 15, 28, 31, 39])
dataframe.to_excel(jobPath + "\\output.xlsx")

# ============================== Format the PDF file ==========================================

# Slice the PDF each page. 
# Read in output file from Export-TCS
# countiesList = for(i in range 1:4000) if AC[i] != countiesList then countiesList = countiesList + AC[i]
# The above line is adding all the counties to a countiesList
# We're going to use that list, along with the installment period and year to name the PDF files
# If the number of sliced PDF's - 1 (for the ending) matches the number of agencies that we have to pay, continue, otherwise
# Throw an error and say "Number of PDF's does not match number of agencies that we have to pay.
# Please manually split, and name the PDFs is not done so. Then, press 6 to run the next step in the processing tax payment process

# ============================ Run the loans through JTT ======================================

# Use the same call that JTT is using to call the SQL database and save an excel sheet in the proper folder

# ==================== JTT Parsing, output Do Not Pay and AFR Analysis ========================

### COMPLETED - NEEDS INTEGRATION ###  Read in the JTT file, output a file with the Do Not Pay, and AFR analysis. 
### COMPLETED - NEEDS INTEGRATION ###  Open Excel File
# If loop through each row, and if the AG == "Y" && AI == "Y" then add that row's columns to Do Pay List.xlsx
# Open the Do Not Pay central file in /TAXES/TAX_TOOLS/Do_Not_Pay.xlsx and 
# Add the Do Not Pay entry 
# Save the Do Not Pay workbook
# Close the Do Not Pay File
# If loop through each row, and if the AG == "N" && AI != "N" then add that row's columns 38, 39, 40, 41 to a Do_Not_Pay.xlsx
# Open the AFR analysis central file in Shared/Escrow_Administration
# Add the AFR analysis entry
# Save the AFR Analysis workbook
# Close the AFR Analysis workbook
# If loop through each row, and if the AI column.value == "Y" then add that row's columns 37, 38, 39 to a AFR_analysis_sheet.xlsx
# Create an Excel file for the Do_Pay_List.xlsx 


# ======================== Input Intact Info, Create IPT ======================================

# Open Intacct.xlsx
# Open Export-TS36501 or the Export-TS36501 variant
# Copy the values from Export-TS36501 and put them into the proper location in Intacct.xlsx
# Press Create IPT
# Give it the name IPT
# Close Export-TS36501 with no need to save
# Save Intact.xlsx
# Close Intacct.xlsx

# ======================== Add the bills to the borrowers folders =============================

# Open the tool used to add the bills to the borrowers folders
# Look in the local directory
# Copy the PDF names and paste them into the proper line
# Open Export-TS36501 and count the number of loans - 1 (for header)
# Press create