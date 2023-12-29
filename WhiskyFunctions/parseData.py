# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:05:53 2022

@author: grant
"""

def parseData(Data):
    NewData = []
    
    for i in Data:
        if i[2] == "":
            
            
            add1 = i[0]
            add2 = i[1]
            
            add3 = add1.replace(" ", "")
            add4 = add2.replace(" ", "")
            
            add = add3 + " " + add4
            
            NewData.append([add, i[3], i[4], i[5], i[6]])
        else:
            NewData.append([i[2], i[3], i[4], i[5], i[6]])
    return(NewData)