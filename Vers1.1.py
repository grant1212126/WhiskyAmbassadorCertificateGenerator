# -*- coding: utf-8 -*-
"""
Created on Wed Nov 25 01:00:07 2020

@author: grant
"""

import os
import tkinter as tk
from shutil import copy

from WhiskyFunctions.addInfoCertificate import addInfoCertificate
from WhiskyFunctions.addInfoHistory import addInfoHistory
from WhiskyFunctions.parseData import parseData
from WhiskyFunctions.parseError import parseError
from WhiskyFunctions.searchFile import SearchFile
from WhiskyFunctions.searchFileHistory import SearchFileHistory
from WhiskyFunctions.processConfigString import processConfigString

root = tk.Tk()

root.title("Whisky Ambassador Program")

windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()

positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

root.geometry("+{}+{}".format(positionRight, positionDown))

lbl = tk.Label(root, text="Please select from one of the following options")
lbl.grid(column=1, row=0, pady=20, columnspan=2)

closeBtn = tk.Button(root, text="Close", command=root.destroy)
closeBtn.grid(column=3, row=4, padx=5, pady=5)

lblFname = tk.Label(root, text="Generate client award  history certificate")
lblFname.grid(column=1, row=1, padx=25, pady=10)

lblSname = tk.Label(root, text="Generate specific award certificate")
lblSname.grid(column=2, row=1, pady=10)

global CertificateTemplateList, AwardCertificateTemplate, SpreadSheet, SpreadSheetFileExists

SpreadSheetExists = False

if not os.path.exists("config.txt"):
    with open("config.txt", "w+") as f:

        tk.messagebox.showinfo(title="Information", message="Configuration file does not exist, one will be created")


else:

    AwardCertificateTemplate = ""

    with open("config.txt", "r+") as f:
        config = f.readlines()

        CertificateTemplateList = []

        SpreadSheetFileExists = False

        for i in config:
            case = i.strip("\n")

            if "Certificate_template" in case:
                newTemplate = processConfigString(case)

                CertificateTemplateList.append(newTemplate)

                print("Template = " + newTemplate)
            if "Award_history_template" in case:
                AwardCertificateTemplate = processConfigString(case)

                print("Award Certificate Template = " + AwardCertificateTemplate)

            if "Spreadsheet" in case:
                SpreadSheet = processConfigString(case)

                if len(SpreadSheet) > 1:
                    SpreadSheetExists = True

                print("Spreadsheet file = " + SpreadSheet)

        if SpreadSheetExists:
            SpreadSheetFileExists = os.path.isfile("Spreadsheet/" + SpreadSheet)

        if not SpreadSheetExists or not os.path.isfile("Spreadsheet/" + SpreadSheet):
            tk.messagebox.showinfo(title="Information",
                                   message="No Spreadsheet file has been selected to search through or has been "
                                           "deleted, please use the 'Select Spreadsheet' button in the bottom left "
                                           "corner to select a spreadsheet to use")

if not os.path.isdir("Templates/"):  # Checks if template directory exists, if not then it creates it
    os.makedirs("Templates/")
    os.makedirs("Templates/Award History Template/")
    os.makedirs("Templates/Certificate Templates/")
    tk.messagebox.showinfo(title="Information",
                           message="Existing templates folder has not been found, one will be created")

if not os.path.isdir("Spreadsheet/"):
    os.makedirs("Spreadsheet/")
    tk.messagebox.showinfo(title="Information",
                           message="Existing Databse folder has not been found, one will be created")


def changeSpreadsheet():
    global SpreadSheet, SpreadSheetExists, AwardCertificateTemplate, CertificateTemplateList, SpreadSheetFileExists

    tk.Tk().withdraw()
    newSpreadSheet = tk.askopenfilename()

    newSpreadSheet = os.path.basename(newSpreadSheet)

    if ".xlsx" not in newSpreadSheet:
        tk.messagebox.showerror(title="Information",
                                message="Error, you have not selected a valid spreadsheet file, old spreadsheet file "
                                        "(if it exists) will not be replaced, make sure you select a spreadsheed file "
                                        "with the extension '.xlsx'")
        return

    if SpreadSheetFileExists:
        os.remove("Spreadsheet/" + SpreadSheet)
        print("Existing spreadsheet removed")

    with open("config.txt", "w") as f:

        for entry in CertificateTemplateList:
            f.write("Certificate_template=" + entry)
            f.write("\n")

        f.write("Spreadsheet=" + newSpreadSheet + "\n")

    newFile = newSpreadSheet.replace("/", '\\')

    dst = os.getcwd()

    dst = dst + "\\Spreadsheet\\"

    copy(newFile, dst)

    SpreadSheet = newSpreadSheet

    SpreadSheetExists = True

    tk.messagebox.showinfo(title="Information", message="Spreadsheet file has been successfully selected and changed!")


spreadsheetBtn = tk.Button(root, text="Select Spreadsheet", command=lambda: changeSpreadsheet())
spreadsheetBtn.grid(column=0, row=4, padx=10, pady=2)


def createIndividualSearchWindow():
    global SpreadSheetExists

    if not SpreadSheetExists:
        tk.messagebox.showinfo(title="Information",
                               message="A valid spreadsheet file does not exist, please use the 'Select Spreadsheet' "
                                       "button to select a spreadsheet file to use")
        return

    root.wm_state('iconic')

    window3 = tk.Toplevel(root)
    window3.grab_set()

    window3.title("Individual Award Search")

    windowWidth = window3.winfo_reqwidth()
    windowHeight = window3.winfo_reqheight()

    positionRight = int(window3.winfo_screenwidth() / 2 - windowWidth / 2)
    positionDown = int(window3.winfo_screenheight() / 2 - windowHeight / 2)

    window3.geometry("+{}+{}".format(positionRight, positionDown))

    window3.resizable(True, True)

    lbl = tk.Label(window3, text="Enter the name of the person you want to create the certificate for")
    lbl.grid(column=0, row=0, columnspan=2, padx=15, pady=10)

    closeBtn = tk.Button(window3, text="Close", command=window3.destroy)
    closeBtn.grid(column=2, row=4, padx=5, pady=5)

    lblFname = tk.Label(window3, text="First Name")
    lblFname.grid(column=0, row=1, pady=5)

    txt = tk.Entry(window3, width=10)
    txt.grid(column=0, row=2, pady=5)

    lblSname = tk.Label(window3, text="Second Name")
    lblSname.grid(column=1, row=1, pady=5)

    txt2 = tk.Entry(window3, width=10)
    txt2.grid(column=1, row=2, pady=5)

    def createGenerateWindow():

        global chosenTemplate, SpreadSheet

        chosenTemplate = None

        firstName = txt.get().strip()
        secondName = txt2.get().strip()

        output = SearchFile(firstName.lower(), secondName.lower(), FileName=SpreadSheet)

        if type(output) == int:
            tk.messagebox.showinfo("Error information", parseError(output))
        else:

            output = output[0]

            Window2 = tk.Toplevel(root)
            Window2.grab_set()

            Window2.title("Generate Window")

            Window2.resizable(True, True)

            windowWidth = Window2.winfo_reqwidth()
            windowHeight = Window2.winfo_reqheight()

            positionRight = int(Window2.winfo_screenwidth() / 2 - windowWidth / 2)
            positionDown = int(Window2.winfo_screenheight() / 2 - windowHeight / 2)

            Window2.geometry("+{}+{}".format(positionRight, positionDown))

            Window2.update_idletasks()

            combo2 = tk.Combobox(Window2)
            newOutput = []

            for i in output:
                newOutput.append(str(i[0]) + ", " + str(i[1]) + ", " + str(i[2]) + ", " + str(i[3]))

            combo2["values"] = newOutput
            combo2["width"] = 60
            combo2.grid(column=1, row=1, pady=10, padx=10)

            lbl2 = tk.Label(Window2, text="Please choose from one of the options below")
            lbl2.grid(column=1, row=0, padx=10, pady=10)

            def generate():

                if os.path.isdir("Output/") == False:  # Checks if output directory exists, if not then it creates it
                    os.makedirs("Output/")
                    tk.messagebox.showinfo(title="Information",
                                           message="Existing Output folder has not been found, one will be created")

                if chosenTemplate != None:

                    requestedItem = combo2.get()
                    lbl2.configure(text=requestedItem)

                    message = addInfoCertificate(requestedItem, chosenTemplate)
                    tk.messagebox.showinfo("Whiskey Ambassador", message)

                else:
                    tk.messagebox.showerror(title="Error",
                                            message="Error, no template has been selected for certificate generation, please click the 'Templates' button to select one")

            btn2 = tk.Button(Window2, text="Generate", command=lambda: generate())
            btn2.grid(column=1, row=2, pady=10)

            def templates():

                Window4 = tk.Toplevel(root)
                Window4.grab_set()

                Window4.title("Template Window")

                Window4.resizable(True, True)

                windowWidth = Window4.winfo_reqwidth()
                windowHeight = Window4.winfo_reqheight()

                positionRight = int(Window2.winfo_screenwidth() / 2 - windowWidth / 2)
                positionDown = int(Window2.winfo_screenheight() / 2 - windowHeight / 2)

                Window4.geometry("+{}+{}".format(positionRight, positionDown))

                Window4.update_idletasks()

                lbl1 = tk.Label(Window4, text="Please select a template to use or delete by using the selector box")
                lbl1.grid(column=1, row=0, padx=10, pady=10)

                lbl2 = tk.Label(Window4, text="No template has been selected")
                lbl2.grid(column=1, row=5, padx=10, pady=10)

                TemplatesCombo = tk.Combobox(Window4)

                TemplatesCombo["values"] = CertificateTemplateList
                TemplatesCombo["width"] = 30
                TemplatesCombo.grid(column=1, row=1, pady=10, padx=10)

                def selectTemplate():

                    global chosenTemplate

                    requestedItem = TemplatesCombo.get()
                    chosenTemplate = requestedItem
                    if requestedItem == "":
                        tk.messagebox.showinfo("Whiskey Ambassador", "No template has been selected")
                        lbl2.configure(text="No template selected")
                    else:
                        lbl2.configure(text=requestedItem + " has been selected")
                        tk.messagebox.showinfo("Whiskey Ambassador",
                                               "Template has been selected, you can now close this window and return to the previous screen")

                btn1 = tk.Button(Window4, text="Select Template", command=lambda: selectTemplate())
                btn1.grid(column=1, row=2, pady=10)

                def deleteTemplate():

                    requestedItem = TemplatesCombo.get()

                    if requestedItem == "":
                        tk.messagebox.showerror("Error",
                                                "No template has been selected for deletion, please select one using the drop-down menu")
                        return

                    MsgBox = tk.messagebox.askquestion("Template Deletion",
                                                       "Are you sure you would like to delete the currently selected template?")

                    if MsgBox == "yes":

                        global CertificateTemplateList, AwardCertificateTemplate, SpreadSheet

                        with open("config.txt", "w") as f:

                            items = len(CertificateTemplateList)

                            for line in CertificateTemplateList:
                                if line != requestedItem:
                                    f.write("Certificate_template=" + line + "\n")

                            f.write("Award_history_template=" + AwardCertificateTemplate + "\n")
                            f.write("Spreadsheet=" + SpreadSheet + "\n")

                            CertificateTemplateList.remove(requestedItem)

                            curdir = os.getcwd()

                            if os.path.exists(curdir + "\\Templates\\Certificate Templates\\" + requestedItem):
                                os.remove(curdir + "\\Templates\\Certificate Templates\\" + requestedItem)

                            tk.messagebox.showinfo("Whiskey Ambassador", "Template has been deleted")

                            TemplatesCombo["values"] = CertificateTemplateList

                btn2 = tk.Button(Window4, text="Delete Template", command=lambda: deleteTemplate())
                btn2.grid(column=1, row=3, pady=10)

                def addTemplate():

                    global validFile

                    validFile = None

                    newTemplateWindow = tk.Toplevel(root)
                    newTemplateWindow.grab_set()

                    newTemplateWindow.title("Template Window")

                    newTemplateWindow.resizable(True, True)

                    windowWidth = newTemplateWindow.winfo_reqwidth()
                    windowHeight = newTemplateWindow.winfo_reqheight()

                    positionRight = int(newTemplateWindow.winfo_screenwidth() / 2 - windowWidth / 2)
                    positionDown = int(newTemplateWindow.winfo_screenheight() / 2 - windowHeight / 2)

                    newTemplateWindow.geometry("+{}+{}".format(positionRight, positionDown))

                    newTemplateWindow.update_idletasks()

                    lbl1 = tk.Label(newTemplateWindow, text="Create a new template")
                    lbl1.grid(column=1, row=0, padx=10, pady=10, columnspan=2)

                    def newTemplate():

                        global file
                        global validFile

                        tk.Tk().withdraw()
                        file = tk.askopenfilename()

                        if ".docx" not in file:
                            validFile = False
                            tk.messagebox.showinfo(title="Information",
                                                   message="Error, you have not selected a word document file")
                        else:
                            tk.messagebox.showinfo(title="Information", message="File successfuly selected")
                            validFile = True

                    btn1 = tk.Button(newTemplateWindow, text="Select Template File", command=lambda: newTemplate())
                    btn1.grid(column=2, row=1, padx=30)

                    def showAddInfo():
                        tk.messagebox.showinfo(title="Information",
                                               message="Select the word document you want to add as a template, then once selected type the name of the new template into the entry box. Once completed click the 'Add Template' button and your new template will be added")

                    infoBtn = tk.Button(newTemplateWindow, text="Info", command=lambda: showAddInfo())
                    infoBtn.grid(column=0, row=5, padx=5, pady=5)

                    def add():

                        global CertificateTemplateList, AwardCertificateTemplate, SpreadSheet

                        if validFile == None:
                            tk.messagebox.showinfo(title="Information",
                                                   message="Error, you have not selected a file for the template, please click the 'Select Template File' button and select the file you want to use as the template")
                        else:
                            if validFile == False:
                                tk.messagebox.showinfo(title="Information",
                                                       message="Error, you have not selected a file for the template, please click the 'Select Template File' button and select the file you want to use as the template")
                            else:
                                newFile = file.replace("/", '\\')

                                dst = os.getcwd()

                                dst = dst + "\Templates\Certificate Templates"

                                assert os.path.isdir(dst)

                                copy(newFile, dst)

                                with open("config.txt", "w") as f:

                                    newEntry = os.path.basename(newFile)
                                    CertificateTemplateList.append(newEntry)
                                    for entry in CertificateTemplateList:
                                        f.write("Certificate_template=" + entry)
                                        f.write("\n")

                                    f.write("Award_history_template=" + AwardCertificateTemplate + "\n")
                                    f.write("Spreadsheet=" + SpreadSheet + "\n")

                                    TemplatesCombo["values"] = CertificateTemplateList

                                tk.messagebox.showinfo(title="Information",
                                                       message="The template has successfuly been added")

                    addBtn = tk.Button(newTemplateWindow, text="Add Template", command=lambda: add())
                    addBtn.grid(column=2, row=4, padx=10, pady=10)

                    closeBtn = tk.Button(newTemplateWindow, text="Close", command=newTemplateWindow.destroy)
                    closeBtn.grid(column=4, row=5, pady=5, padx=5)

                btn2 = tk.Button(Window4, text="Add new template", command=lambda: addTemplate())
                btn2.grid(column=1, row=4, pady=10)

                closeBtn = tk.Button(Window4, text="Close", command=Window4.destroy)
                closeBtn.grid(column=2, row=5, pady=5, padx=5)

                def showInfo():
                    tk.messagebox.showinfo(title="Information",
                                           message="Use the selector box to select an existing template, then once a template hase been selected, click the 'Select Template' button where you can then close the templates window and return to the previous screen to generate the certificate.")

                infoBtn = tk.Button(Window4, text="Info", command=showInfo)
                infoBtn.grid(column=0, row=5, pady=5, padx=5)

                if len(CertificateTemplateList) == 0:
                    tk.messagebox.showinfo(title="Information",
                                           message="No existing templates were found in the config file, you can add a new template by clicking the add button")

            btn3 = tk.Button(Window2, text="Templates", command=lambda: templates())
            btn3.grid(column=0, row=4, pady=5, padx=5)

            closeBtn2 = tk.Button(Window2, text="Close", command=Window2.destroy)
            closeBtn2.grid(column=2, row=4, pady=5, padx=5)

    btn = tk.Button(window3, text="Search", command=createGenerateWindow)
    btn.grid(column=0, row=3, columnspan=2)


def createAwardHistoryWindow():
    global SpreadSheetExists

    if SpreadSheetExists == False:
        tk.messagebox.showinfo(title="Information",
                               message="A valid spreadsheet file does not exist, please use the 'Select Spreadsheet' button to select a spreadsheet file to use")
        return

    root.wm_state('iconic')

    AwardHistoryWindow = tk.Toplevel(root)
    AwardHistoryWindow.grab_set()

    AwardHistoryWindow.title("Award History Search")

    windowWidth = AwardHistoryWindow.winfo_reqwidth()
    windowHeight = AwardHistoryWindow.winfo_reqheight()

    positionRight = int(AwardHistoryWindow.winfo_screenwidth() / 2 - windowWidth / 2)
    positionDown = int(AwardHistoryWindow.winfo_screenheight() / 2 - windowHeight / 2)

    AwardHistoryWindow.geometry("+{}+{}".format(positionRight, positionDown))

    AwardHistoryWindow.resizable(True, True)

    lbl = tk.Label(AwardHistoryWindow, text="Enter the name of the person you want to create the certificate for")
    lbl.grid(column=0, row=0, columnspan=2, padx=15, pady=10)

    closeBtn = tk.Button(AwardHistoryWindow, text="Close", command=AwardHistoryWindow.destroy)
    closeBtn.grid(column=2, row=4, padx=5, pady=5)

    lblFname = tk.Label(AwardHistoryWindow, text="First Name")
    lblFname.grid(column=0, row=1, pady=5)

    txt = tk.Entry(AwardHistoryWindow, width=10)
    txt.grid(column=0, row=2, pady=5)

    lblSname = tk.Label(AwardHistoryWindow, text="Second Name")
    lblSname.grid(column=1, row=1, pady=5)

    txt2 = tk.Entry(AwardHistoryWindow, width=10)
    txt2.grid(column=1, row=2, pady=5)

    def createGenerateWindow2():

        global SpreadSheet, AwardCertificateTemplate

        chosenTemplate = None

        firstName = txt.get().strip()
        secondName = txt2.get().strip()

        output = SearchFile(firstName.lower(), secondName.lower(), FileName=SpreadSheet)

        if type(output) == int:
            tk.messagebox.showinfo("Error information", parseError(output))
        else:
            Data = output[1]

            output = output[0]

            SearchWindow = tk.Toplevel(root)
            SearchWindow.grab_set()

            SearchWindow.title("Generate Window")

            SearchWindow.resizable(True, True)

            windowWidth = SearchWindow.winfo_reqwidth()
            windowHeight = SearchWindow.winfo_reqheight()

            positionRight = int(SearchWindow.winfo_screenwidth() / 2 - windowWidth / 2)
            positionDown = int(SearchWindow.winfo_screenheight() / 2 - windowHeight / 2)

            SearchWindow.geometry("+{}+{}".format(positionRight, positionDown))

            SearchWindow.update_idletasks()

            combo2 = tk.Combobox(SearchWindow)
            newOutput = []

            for i in output:
                newOutput.append(str(i[0]) + ", " + str(i[1]) + ", " + str(i[2]) + ", " + str(i[3]))

            combo2["values"] = newOutput
            combo2["width"] = 60
            combo2.grid(column=1, row=1, pady=10, padx=10)

            lbl2 = tk.Label(SearchWindow, text="Please choose from one of the options below")
            lbl2.grid(column=1, row=0, padx=10, pady=10)

            def generateHistory():

                global AwardCertificateTemplate

                if os.path.isdir("Output/") == False:  # Checks if output directory exists, if not then it creates it
                    os.makedirs("Output/")
                    tk.messagebox.showinfo(title="Information",
                                           message="Existing Output folder has not been found, one will be created")

                if AwardCertificateTemplate != None:

                    requestedItem = combo2.get()

                    requestedList = requestedItem.split(",")

                    certId = requestedList[1]

                    if certId == "No ID number":
                        message = parseError(4)
                        tk.messagebox.showerror(title="Error", message=message)

                    else:
                        output = SearchFileHistory(certId, Data)

                        lbl2.configure(text=requestedItem)

                        message = addInfoHistory(output, AwardCertificateTemplate)
                        tk.messagebox.showinfo("Whiskey Ambassador", message)

                else:
                    tk.messagebox.showerror(title="Error",
                                            message="Error, no template exists for award history generation, please click the 'Templates' button to add one")

            btn2 = tk.Button(SearchWindow, text="Generate Award History", command=lambda: generateHistory())
            btn2.grid(column=1, row=2, pady=10)

            def historyTemplate():

                global CertificateTemplateList, AwardCertificateTemplate, SpreadSheet

                if AwardCertificateTemplate != None:
                    MsgBox = tk.messagebox.askquestion(title="Information",
                                                       message="A template already exists for award generation, would you like to replace it?")

                    if MsgBox == "no":
                        return

                else:
                    tk.messagebox.showinfo(title="Information",
                                           message="No existing template exists, please select a .docx file to use as an award history template")

                tk.Tk().withdraw()
                file = tk.askopenfilename()

                if ".docx" not in file:
                    tk.messagebox.showinfo(title="Information",
                                           message="Error, you have not selected a word document file")
                    return
                else:

                    # Adds new Template to directory

                    newFile = file.replace("/", '\\')

                    dst = os.getcwd()

                    dst = dst + "\Templates\Award History Template"

                    assert os.path.isdir(dst)

                    copy(newFile, dst)

                    if AwardCertificateTemplate != None:
                        curdir = os.getcwd()

                        os.remove(curdir + "\\Templates\\Award History Template\\" + AwardCertificateTemplate)

                    with open("config.txt", "w") as f:

                        AwardCertificateTemplate = os.path.basename(newFile)

                        for entry in CertificateTemplateList:
                            f.write("Certificate_template=" + entry)
                            f.write("\n")

                        f.write("Award_history_template=" + AwardCertificateTemplate + "\n")
                        f.write("Spreadsheet=" + SpreadSheet + "\n")

                    tk.messagebox.showinfo(title="Information", message="The template has successfuly been changed")

            btn3 = tk.Button(SearchWindow, text="Templates", command=lambda: historyTemplate())
            btn3.grid(column=0, row=4, pady=5, padx=5)

            closeBtn2 = tk.Button(SearchWindow, text="Close", command=SearchWindow.destroy)
            closeBtn2.grid(column=2, row=4, pady=5, padx=5)

    btn = tk.Button(AwardHistoryWindow, text="Search", command=lambda: createGenerateWindow2())
    btn.grid(column=0, row=3, columnspan=2)


btn1 = tk.Button(root, text="Start", command=createAwardHistoryWindow)
btn1.grid(column=1, row=2)

btn2 = tk.Button(root, text="Start", command=createIndividualSearchWindow)
btn2.grid(column=2, row=2)

root.mainloop()
