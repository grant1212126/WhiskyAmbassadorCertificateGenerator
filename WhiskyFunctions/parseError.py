# -*- coding: utf-8 -*-
"""
Created on Sun Jul 24 21:06:22 2022

@author: grant
"""

def parseError(errorcode):
    if errorcode == 1:
        return("Inputed name is not in database")
    if errorcode == 2:
        return("Error with adding information to Certificate template, check that the <Insert -------> tags are where you want the item to be inserted")
    if errorcode == 3:
        return("Error, you have not entered either a first or second name to search, please enter either or both a second and first name")
    if errorcode == 4:
        return("Error, the person selected does not have a certificate ID number, please check the database to see if this person has an ID number")
    else:
        return("Unkown error")