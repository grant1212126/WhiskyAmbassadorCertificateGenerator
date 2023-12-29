# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:07:00 2022

@author: grant
"""

import os
from docx import Document

from WhiskyFunctions.parseError import parseError


def addInfoCertificate(Data, Template):
    
    inserted = False
    
    temp = os.getcwd()
    
    temp = temp + "\\Templates\\Certificate Templates\\" + Template
    
    doc = Document(temp)

    paragraphs = doc.paragraphs
    
    Parse = Data.split(", ")
    
    
    FullName = Parse[0]
    ID = Parse[1]
    Date = Parse[2]
    
    for para in paragraphs:
        items = para.runs
        for line in items:
            print("line: " + line.text)
        if "<Insert Name>" in para.text:
            items = para.runs
            for line in items:
                if "<Insert Name>" in line.text:
                    line.text = str(FullName)
            inserted = True
            
        if "<Insert Certificate Number>" in para.text:
            items = para.runs
            for line in items:
                if "<Insert Certificate Number>" in line.text:
                    line.text = "Cert No: " + str(ID)

            inserted = True
        
        if "<Insert Issue Date>" in para.text:
            items = para.runs
            for line in items:
                if "<Insert Issue Date>" in line.text:
                    line.text = str(Date)
            
            inserted = True
    
    if inserted == False:
        return(parseError(2))
    
    try:
        doc.save(os.getcwd() + "\\Output\\" + str(FullName) +  ".docx")
    except Exception as e:
        return("There was an error with Certificate generation, do you have the certificate of " + str(FullName) + " already open?")
    
    return("Certificate creation successful, please see " + str(FullName) + ".docx for completed certificate")
