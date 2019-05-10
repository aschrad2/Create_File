# -*- coding: utf-8 -*-
"""
Created on Fri May 10 14:51:59 2019

Add in the support files to the newly created folder

@author: austin.schrader
"""

import os
import shutil

sourceBillExport = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\Bill export B Tool.xlsm"
sourceDoNotPay = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\For Russell to review-Loan Template.xlsm"
sourceForRussell = "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\ATemplateCopyFolder\\INTACCT Escrow Upload Template.xlsm"

shutil.copyfile(sourceBillExport, "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Destination\\Bill export B Tool.xlsm")
shutil.copyfile(sourceDoNotPay, "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Destination\\For Russell to review-Loan Template.xlsm")
shutil.copyfile(sourceForRussell, "C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Destination\\INTACCT Escrow Upload Template.xlsm")

