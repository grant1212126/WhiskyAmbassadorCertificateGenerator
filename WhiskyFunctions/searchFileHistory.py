# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:14:45 2022

@author: grant
"""

from WhiskyFunctions.parseData import parseData

def SearchFileHistory(certNum, Data):
    
    Matches = []
    
    print(certNum)
    
    for entry in Data:
        if certNum.strip() in entry[3].strip(): # [3] corresponds to where the certificate number is stored in the Data variable
            Matches.append(entry)
        
    return(parseData(Matches))