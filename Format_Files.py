# -*- coding: utf-8 -*-
"""
Created on Tue May 14 09:51:29 2019

@author: austin.schrader
"""

# To use the openpyxl module, we first have to convert the Excel file Export-TCS36501 from xls into xlsx file format
import win32com.client as win32
# Openpyxl is the module that's doing the editing of an excel. However, it can only wrok with .xlsx formats
#import openpyxl
import pandas

# Job fileName and job stateCode, hardcoded which should be replaced later
fileName = "\\NJ666666"
stateCode = "\\NJ"

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
