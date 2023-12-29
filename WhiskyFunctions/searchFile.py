# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:03:46 2022

@author: grant
"""

import xlrd
import os

from WhiskyFunctions.parseData import parseData

def SearchFile(firstName, secondName, FileName):
    
    if secondName == "":
        secondName = "------"
        
    if firstName == "":
        firstName = "------"
        
    if secondName == "------" and firstName == "------":
        errorCode = 3
        return errorCode; # Returns integer error code if user did not enter a value for both second and first name 

    Data = []
    
    Matches = []
    
    errorCode = 0
    
    spreadSheet = xlrd.open_workbook(os.getcwd() + "\\Spreadsheet\\" + FileName)
    
    sheet = spreadSheet.sheet_by_name("Delegate info")
    
    for row in range(sheet.nrows):
            row_value = sheet.row_values(row)
            if str(row_value[7]) != "" and type(row_value[7]) == float and row_value[7] < 99999.0:
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row_value[7], spreadSheet.datemode)
                if row_value[31] == "":    
                    Data.append([str(row_value[2]).strip(), str(row_value[3]).strip(), str(row_value[4]).strip() , "No ID number", "{0}/{1}/{2}".format(day, month, year), str(row_value[6]).strip(), str(row_value[27]).strip()])
                else:
                    Data.append([str(row_value[2]).strip() , str(row_value[3]).strip(), str(row_value[4]).strip(), str(row_value[31]).strip(), "{0}/{1}/{2}".format(day, month, year), str(row_value[6]).strip(), str(row_value[27]).strip()])
            else:
                if row_value[31] == "":    
                    Data.append([str(row_value[2].strip()), str(row_value[3]).strip(), str(row_value[4]).strip(), "No ID number", "Error receiving date", str(row_value[6]).strip(), str(row_value[27]).strip()])
                else:
                    Data.append([str(row_value[2]).strip(), str(row_value[3]).strip(), str(row_value[4]).strip(), str(row_value[31]).strip(), "Error receiving date", str(row_value[6]).strip(), str(row_value[27]).strip()])
    
    # Verbos search, searches through and returns each reccord containing EITHER first or second name, used when user doesent specify either a first or second name
            
    if firstName == "------" or secondName == "------":
        for entry in Data:
            if firstName in entry[0].lower() or firstName in entry[2].lower():
                Matches.append(entry)
                
            elif secondName in entry[1].lower() or secondName in entry[2].lower():
                Matches.append(entry)
    else:
        for entry in Data:
            if firstName in entry[0].lower() or firstName in entry[2].lower():
                if secondName in entry[1].lower() or secondName in entry[2].lower():
                    Matches.append(entry)
    
    if Matches == []:
        errorCode = 1
    
    if errorCode == 0:
        output = []
    
        for i in Matches:
            output.append(i)
        
        return(parseData(output), Data)
    
    else:
        return(errorCode)