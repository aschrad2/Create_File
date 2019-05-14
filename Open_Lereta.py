# -*- coding: utf-8 -*-
"""
Created on Mon May 13 10:22:19 2019

@author: austin.schrader
"""

import os.path
import win32com.client

input("Press Enter to continue once you've downloaded the Export, and PDF to the file...")

if os.path.exists("excelsheet.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(os.path.abspath("excelsheet.xlsm"), ReadOnly=1)
    xl.Application.Run("excelsheet.xlsm!modulename.macroname")
##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl