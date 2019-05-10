# -*- coding: utf-8 -*-
"""
Created on Fri May 10 13:04:05 2019

The purpose of the file to complete the first step of paying property tax amounts.

More specifically the first step is to create a job file in the proper folder. 
NJ181818 would be created at the file within CTA Paid Taxes

@author: austin.schrader
"""

import os

# Ask for the job number that the payer would like to start processing for payment
jobRequested = input("What property tax job would you like to process?")

# Add in the necessary backslashes to a) add a backslash, 2) provide the escape. Ex: \NJ484848
fileName = "\\" + jobRequested

# Read the jobRequested's stateCode and append that state to the path
stateCode = "\\" + fileName[1:3]
print(stateCode)

# Append the inputted fileName 
path = "\\\\cottonwood\\Users\\Shared\\Taxes\\CTA Paid Taxes\\2019\\" + stateCode + fileName

# Make the
try:
    os.mkdir(path)
except OSError:
    print("Creation of the directory %s failed" % path)
else:
    print("Succesfully created the directory %s " % path)