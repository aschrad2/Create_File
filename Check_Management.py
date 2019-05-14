# -*- coding: utf-8 -*-
"""
Created on Tue May 14 11:07:55 2019

JTT Parsing, output Do Not Pay and AFR Analysis 

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

@author: austin.schrader
"""
import pandas

# Job fileName and job stateCode, hardcoded which should be replaced later
fileName = "\\NJ666666"
stateCode = "\\NJ"

# Establishes the path for the .xlsx file (the version that we can work with)
jobPath = ("\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName)
exportFilePath = "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019" + stateCode + fileName + "\\For Russell to review-Loan Template.xlsm"

# Reads the excel and parses out columns 0,1,4,6 etc (all the ones we need.)
# Then, it exports the file to the jobPath + \\output.xlsx
dataframe = pandas.read_excel(exportFilePath, parse_cols = [0,1,2,3,4,23,32,34,36])
dataframe.to_excel(jobPath + "\\managementoutput.xlsx")

# Creates a specific dataframe that contains the records of loans that need management approval. This can be added
# to the do not pay, seek management approval section of the Do_Not_Pay.xlsx spreadsheet next
mgmtDataFrame = dataframe[dataframe["SentToMgmtFlag"] == "Y"]
mgmtDataFrame.to_excel(jobPath + "\\addtomgmt.xlsx")
print(mgmtDataFrame)

# Creates a specific dataframe that contains the records of loans that will not be paid today. The reason include delinquency on the loan
# In which case, they should be deleted from the Export-TCS36501 file
payDataFrame = dataframe[dataframe["PayFlag"] == "N"]
payDataFrame.to_excel(jobPath + "\\addtodonotpay.xlsx")
print(payDataFrame)

# Creates a specific dataframe that contains the records of loans that need EA analysis. Thus, these loans should be 
# added to the EA Analysis spreadsheet
eaDataFrame = dataframe[dataframe["EAFlag"] == "Y"]
eaDataFrame.to_excel(jobPath + "\\addtoea.xlsx")
print(eaDataFrame)



# =============================================================================
# dataframe.to_excel(jobPath + "\\donotpay.xlsx")
# dataframe.to_excel(jobPath + "\\addtoafr.xlsx")
# dataframe.to_excel(jobPath + "\\askmgmt.xlsx")
# =============================================================================

#do not pay = \\cottonwood\Users\Shared\Taxes\TAX_TOOLS




# Read in the JTT file, output a file with the Do Not Pay, and AFR analysis. 
# Open Excel File
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
