import customtkinter
import string
from tkinter import *
from tkinter import filedialog
import os
from importSpreadSheet import *
from slides import *
from errorHandler import *

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

app = customtkinter.CTk()
app.geometry("600x500")
app.title("Excel Transformer")
app.grid_rowconfigure(1, weight=0)
app.grid_columnconfigure(0, weight=3)

frame = customtkinter.CTkLabel(app, text="")
frame.grid(row=0, column=0)

def getValues(ppFileName):
    excelFile = excelFilePathEntry.get()
    row = headerRowEntry.get()
    col = headerColEntry.get()

    excelFile, ppFileName, row, col = validateInput(excelFile, ppFileName, row, col)
    errors = errorHandling(excelFile, ppFileName, row, col)

    if errors:
        print("errors found")
        createErrorWindow()
        # getErrors(errors)
    else:
        powerPresentation(excelFile, ppFileName, row, col)

def powerPresentation(spreadSheet, ppFileName, headerRow, headerCol):
    ws = loadSpreadSheet(spreadSheet)
    categories = getHeaderNames(headerRow, headerCol, ws)
    items = getSlideDetails(categories, ws, headerRow, headerCol)
    presentation = slideShow(ppFileName)

    for item in items:
        getTopicSlide(presentation, item, categories)

    savePP(presentation, ppFileName)

def validateInput(excelFile, powerPName, row, col):
    excelFilePath = isExcelFileValid(excelFile)
    ppFileName = isFileNameValid(powerPName)
    headerRow = isHeadeRowValid(row)
    headerCol = isHeadeColValid(col)

    return excelFilePath, ppFileName, headerRow, headerCol

def errorHandling(excelFile, ppFile, row, col):
    errors = []
    if excelFile == FALSE:
        errors.append("Invalid or missing excel file.")
    if ppFile == FALSE:
        errors.append("Invalid power point name.")
    if row == FALSE:
        errors.append("Invalid header row entry.")
    if col == FALSE:
        errors.append("Invalid header column entry.")
    return errors

def isFileNameValid(fileName):
    return fileName + ".pptx" if isinstance(fileName, str) else FALSE
     
def isExcelFileValid(filepath):
    return filepath if isinstance(filepath, str) else FALSE 

def isHeadeRowValid(row):
    return int(row) if row and isinstance(int(row), int) else FALSE 

def isHeadeColValid(col):
    return colToNumber(col) if col and isinstance(col, str) else FALSE 

def colToNumber(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def setFilePath(filepath):
    excelFilePathEntry.insert(0, filepath)

def selectExcelFile():
    excelFilePath = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select File",
        filetypes=(("excel files", ".xlsx"), ("other files", ".xls"))
    )
    setFilePath(excelFilePath)

def savePowerPoint():
    ppFilePath = filedialog.asksaveasfilename(
        initialdir=os.getcwd(),
        title="Save As"
    )
    getValues(ppFilePath)
    
excelFilePathEntry = customtkinter.CTkEntry(frame, width=200)
headerRowEntry = customtkinter.CTkEntry(frame, placeholder_text="1", width=10)
headerColEntry = customtkinter.CTkEntry(frame, placeholder_text="A", width=10)
excelLabel = customtkinter.CTkLabel(frame, text="Excel File:")
headerRowLabel = customtkinter.CTkLabel(frame, text="File Header Row:")
headerColLabel = customtkinter.CTkLabel(frame, text="File Header Column:")
appTitle = customtkinter.CTkLabel(frame, text="The Excel Transformer")
appSubtitle = customtkinter.CTkLabel(frame, text="Create a powerpoint presentation from your excel spreadsheet.")
selectBtn = customtkinter.CTkButton(frame, text="Browse", width=50, command=selectExcelFile)
submitBtn = customtkinter.CTkButton(frame, text="Submit", width=50, command=savePowerPoint)

appTitle.grid(row=0, column=0, columnspan=3, padx=1, pady=5, ipady=0)
appSubtitle.grid(row=1, column=0, columnspan=3, padx=1, pady=5, ipady=0)
excelLabel.grid(row=2, column=0, padx=1, pady=5, ipady=0, sticky="e")
headerRowLabel.grid(row=3, column=0, padx=1, pady=5, ipady=0, sticky="e")
headerColLabel.grid(row=3, column=2, padx=1, pady=5, ipady=0, sticky="e")
excelFilePathEntry.grid(row=2, column=1, padx=1 , pady=5, ipady=0, sticky="w")
headerRowEntry.grid(row=3, column=1, padx=1 , pady=5, ipady=0, sticky="w")
headerColEntry.grid(row=3, column=3, padx=1 , pady=5, ipady=0, sticky="w")
selectBtn.grid(row=2, column=2, padx=1 , pady=5, ipady=0, sticky="w")
submitBtn.grid(row=5, column=1, padx=1 , pady=5, ipady=0, sticky="w")

app.mainloop()