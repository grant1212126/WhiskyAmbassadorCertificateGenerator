# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:12:52 2022

@author: grant
"""

def processConfigString(Data):
    newTemplate = ""
    for i in reversed(Data):
        if i == "=":
            break
        newTemplate = newTemplate + i
        
    newTemplate = newTemplate [::-1]
                
    return(newTemplate)
